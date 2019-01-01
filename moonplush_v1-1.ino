#include <Wire.h> // must be included here so that Arduino library object file references work
#include <RtcDS3231.h>
#include <avr/sleep.h>

RtcDS3231<TwoWire> Rtc(Wire);

// astrological constants
const double mlm_e = 91.929336;     // mean longitude of the moon at epoch 2010
const double elsp_e = 283.112438;   // ecliptic longitude of the sun's perigee at epoch 2010
const double els_e = 279.557208;    // ecliptic longitude of the sun at epoch 2010
const double eos_e = 0.016705;      // eccentricity of the sun's orbit at epoch 2010
const double je = 2455196.5;        // Julian date/time of epoch 2010 (00:00 December 31 2009)

String cmdIn;                       // Used for storing serial commands
static bool debug = true;           // Used to enable dPrint commands
static double pi = 3.141592653589793;// Pi to 15 decimal places

double rtc2julian(RtcDateTime t)               // Converts Unix DateTime into Julian DateTime
{
   return (t / 86400.0f) + 2440587.5f;
}

double date2julian(RtcDateTime t) {
  uint8_t mon = t.Month();
  uint16_t yr = t.Year();
  if (mon < 3){
    mon += 12;
    yr --;
  }
  
  return 2 - floor(yr/100) + floor(yr/400) + t.Day() + floor(30.6001 * (1 + mon)) + floor(365.25 * (yr + 4716)) - 1524.5;
}

double time2dec(RtcDateTime t) {
  return (double)(t.Hour() / 24.0) + (double)(t.Minute() / (60.0 * 24.0)) + (double)(t.Second() / (60.0 * 60.0 * 24.0));
}

double now2julian(RtcDateTime t) {
  uint8_t mon = t.Month();
  uint16_t yr = t.Year();
  if (mon < 3){
    mon += 12;
    yr --;
  }
  
  double jd = 2 - floor(yr/100) + floor(yr/400) + t.Day() + floor(30.6001 * (1 + mon)) + floor(365.25 * (yr + 4716)) - 2456721;
  return jd + (double)(t.Hour() / 24.0) + (double)(t.Minute() / (60.0 * 24.0)) + (double)(t.Second() / (60.0 * 60.0 * 24.0));
}

double consAngle(double angle) {
  return angle - 360 * ((int32_t)(angle / 360.0));
}

bool updateSerial() {
  if (Serial.available()) {                   // If incomming serial data is waiting
    char charIn = Serial.read();              // Read a single character per loop
    switch (charIn) {                         // Decide what to do by which character was received
      case '\n':                              // Skip new line codes
        break;                                // Break from Switch block

      case '\r':
        Serial.flush();                       // Remove any leftover data from the serial buffer
        return true;                          // Return true for full command received
        break;                                // Not technically necessary but improves readability

      default:                                // For all other characters
        cmdIn += charIn;                      // Add the character to the command string
    }
  }
  return false;                               // Return false for no command received
}

void setup() {
  // put your setup code here, to run once:
  Serial.begin(57600);
  Rtc.Begin();
  RtcDateTime compiled = RtcDateTime(__DATE__, __TIME__);
  Serial.print("format for reference: ");
  Serial.println(compiled);

  if (!Rtc.IsDateTimeValid()) 
  {
      Serial.println("RTC lost confidence in the DateTime!");
      Rtc.SetDateTime(compiled);
  }

  if (!Rtc.GetIsRunning())
  {
      Serial.println("RTC was not actively running, starting now");
      Rtc.SetIsRunning(true);
  }

  RtcDateTime now = Rtc.GetDateTime();
  if (now < compiled) 
  {
      Serial.println("RTC is older than compile time!  (Updating DateTime)");
      Rtc.SetDateTime(compiled);
  }
}

void loop() {
  // put your main code here, to run repeatedly:
  if (updateSerial()) {
    parseSerial();
    cmdIn = "";
  }
    
}

double moonPhase(double t) {
  //double d = t - je;                                               // Adjust for days since 2010 epoch
  double d = t;
  dPrint("days since epoch: ");
  dPrintln(String(d));
  
  double mas = consAngle(360.0/365.242191 * d + els_e - elsp_e);            // Calculate mean anomaly of the sun
  dPrint("Mean anomaly of sun: ");
  dPrintln(String(mas));
  
  double tas = consAngle(mas + 360.0/pi * eos_e * sin(mas * pi / 180.0));   // Calculate the true anomaly of the sun
  dPrint("True anomaly of sun: ");
  dPrintln(String(tas));
  
  double ls = consAngle(tas + elsp_e);                                        // Calculate the longitude of the sun
  dPrint("Longitude of sun: ");
  dPrintln(String(ls));
  
  double lm = consAngle(13.1763966f * d + mlm_e);                             // Calculate the longitude of the moon
  dPrint("Longitude of moon: ");
  dPrintln(String(lm));
  
  double D = consAngle(lm - ls);                                   // Calculate the age of the moon
  if (D < 0)
    D += 360.0f;                                                  // I THINK this fixes it?  Maybe?
  dPrint("Age of moon: ");
  dPrint(String(D));
  if (D < 180)
    dPrintln(" waxing");
  else
    dPrintln(" waning");
  
  return ((1.0 - cos(D * pi / 180.0))/ 2.0);                        // Calculate and return the phase of the moon
}

void dPrint(String msg) {
  if (debug) Serial.print(msg);       // Print debug message if debug enabled
}

void dPrintln(String msg) {
  if (debug) Serial.println(msg);     // Print debug message line if debug enabled
}

void parseSerial() {
  //String cmd = cmdIn.substring(0, 2).toUpperCase();         // Command will always be the first 2 characters received
  String cmd = cmdIn.substring(0, 2);
  cmd.toUpperCase();
  dPrint("Command is ");
  dPrintln(cmd);
  
  /*if (cmd == "SS") {                                        // Check if settings are being saved first
    dPrintln("Save Settings");
    if (settingsUpdated) {
      saveConfig();
      dPrintSettings();
    } else
      dPrintln("No settings have been changed");
    return ;
  }*/

  /*settingsUpdated = true;                                   // Assume settings changed*/

  //String val = cmdIn.substring(3, cmdIn.length()).toLowerCase();
  String val = cmdIn.substring(2, cmdIn.length());
  val.toLowerCase();
  val.trim();
  dPrint("Value is ");
  dPrintln(val);

  if (cmd == "MP") {
    //dPrintln("I'll calculate the moon phase here someday.");
    dPrint("The current moon phase of ");
    if (val.length() == 0) {
      RtcDateTime now = Rtc.GetDateTime();
      //double fNow = date2julian(now) + time2dec(now);
      double fNow = now2julian(now);
      dPrint(String(fNow));
      dPrint(": ");
      dPrintln(String(moonPhase(fNow)));
    }else {
      dPrint(val);
      dPrint(": ");
      dPrintln(String(moonPhase(val.toDouble())));
    }
  }

  else if (cmd == "CA") {
    dPrint("Constrained Angle of ");
    dPrint(val);
    dPrint(" degrees: ");
    dPrintln(String(consAngle(val.toDouble())));
  }

  else if (cmd == "GT") {
    RtcDateTime now = Rtc.GetDateTime();
    String theDate;
    theDate = String(now.Month()) + "/" + String(now.Day()) + "/" + String(now.Year()) + " " + String(now.Hour()) + ":" + String(now.Minute()) + ":" + String(now.Second());
    dPrintln(theDate);
    dPrintln(String(now));
  }

  else if (cmd == "ST") {
    //dPrint("Value is ");
    //dPrintln(String(strTo32(val)));
    //dPrintln(String(val.toInt()));
    Rtc.SetDateTime(val.toInt());

    RtcDateTime now = Rtc.GetDateTime();
    String theDate;
    theDate = String(now.Month()) + "/" + String(now.Day()) + "/" + String(now.Year()) + " " + String(now.Hour()) + ":" + String(now.Minute()) + ":" + String(now.Second());
    dPrintln(theDate);
    dPrintln(String(now));
  }

  else if (cmd == "GJ") {
    double fNow;
    RtcDateTime now = Rtc.GetDateTime();
    fNow = date2julian(now);
    fNow += time2dec(now);
    dPrint("Julian date/time: ");
    dPrintln(String(fNow));
    fNow = date2julian(now);
    dPrint("Julian date: ");
    dPrintln(String(fNow));
    fNow = time2dec(now);
    dPrint("Julian time: ");
    dPrintln(String(fNow));
    dPrintln("---");
    dPrintln(String(now2julian(now)));

    /*for (double x = 0; x < 1; x += 0.1) {
      dPrint(String(x));
      dPrint(", ");
      if ((x * 10) == floor(x * 10))
        dPrintln("");
    }*/
  }

  /*if (cmd == "ST") {
      dPrint("Short Timeout: ");
      storage.timeoutShort = val.toInt();
      dPrint(storage.timeoutShort);
      dPrintln("ms");

  } else if (cmd == "LT") {
      dPrint("Long Timeout: ");
      storage.timeoutLong = strTo32(val);
      dPrint(storage.timeoutLong);
      dPrintln("ms");

  } else if (cmd == "CH") {
      dPrint("Wireless Channel: ");
      storage.channel = (uint8_t)val.toInt();
      dPrintln(storage.channel);

  } else if (cmd == "TX") {
      dPrint("Transmit Address: ");
      cmdIn.substring(3, cmdIn.length()).toCharArray(storage.txAddr, 6);
      dPrintln(storage.txAddr);

  } else if (cmd == "RX") {
      dPrint("Receive Address: ");
      cmdIn.substring(3, cmdIn.length()).toCharArray(storage.rxAddr, 6);
      dPrintln(storage.rxAddr);

  } else if (cmd == "LP") {
      dPrint("Long Press Time: ");
      storage.btnHoldTime = val.toInt();
      dPrint(storage.btnHoldTime);
      dPrintln("ms");

  } else if (cmd == "HF") {
      dPrint("Haptic Feedback: ");
      if ((val == "true") || (val == "on")) {
        dPrintln("ON");
        storage.fbHaptic = true;
      } else {
        dPrintln("OFF");
        storage.fbHaptic = false;
      }
        

  } else if (cmd == "NF") {
      dPrint("NeoPixel Feedback: ");
      if ((val == "true") || (val == "on")) {
        dPrintln("ON");
        storage.fbNeopixel = true;
      } else {
        dPrintln("OFF");
        storage.fbNeopixel = false;
      }

  } else if (cmd == "CF") {
      dPrint("Click Frequency: ");
      storage.clickFreq = val.toInt();
      dPrint(storage.clickFreq);
      dPrintln("Hz");

  } else if (cmd == "CD") {
      dPrint("Click Delay: ");
      storage.clickDelay = val.toInt();
      dPrint(storage.clickDelay);
      dPrintln("ms");

  } else if (cmd == "LS") {
      dPrintln("List Settings");
      dPrintSettings();

  } else if (cmd == "XX") {
      dPrintln("Reload Settings");
      loadConfig();
      dPrintSettings();
      settingsUpdated = false;
      
  }else {
      dPrintln("Unrecognized.");
      settingsUpdated = false;              // Command not recognized, do not save settings
      return;
  }*/
}

uint32_t strTo32(String in) {
  in.trim();
  if (in.length() > 4) {
    uint32_t msb = in.substring(0,5).toInt() * pow(10, max(in.length() - 4, 0));
    uint32_t lsb = in.substring(5, in.length()).toInt();
    /*dPrint("MSB ");
    dPrintln(msb);
    dPrint("LSB ");
    dPrintln(lsb);*/
    return (uint32_t)(msb + lsb);
  } else {
    return (uint32_t)(in.toInt());
  }
}
