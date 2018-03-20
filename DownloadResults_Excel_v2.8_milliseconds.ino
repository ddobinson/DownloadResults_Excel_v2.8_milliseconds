/**
 * 
 * Produces output from PICC to Excel:
 * First name, surname, race number, ss1 time, ss2 time, ss3 time, ss4 time, ss5 time, ss6 time, total time
 * Also produces output on sheet "DownloadResults" to print out to thermal printer
 * 
 */
//LIBRARIES
#include <SPI.h>
#include <Wire.h>
#include <MFRC522.h>
#include <MFRC522Extended.h>  //Library extends MFRC522.h to support RATS for ISO-14443-4 PICC.
#include <LCD5110_Graph.h> 
#include <DS3232RTC.h>
#include <TimeLib.h>
//#include <Streaming.h>      //not sure this is needed? Dan
#include <rExcel.h>
#include "pitches.h"     //if error message involving pitches.h comes up then make sure pitches.h file is in the same directory as the sketch
#include "iTimeAfricaV2.h"
#include <stdlib.h>
#include <EEPROM.h>

//GLOBAL DEFINITIONS
int tagCount = 0;                                           // Cont that incremetns on eatch tag read
int flashLed = LOW;                                         // LED state to implement flashing
int typeIndex = 0xFF;                                       // Index for pod identity in the pod Type List
podType thisPod ;                                           // Identity configured for tehis pod
unsigned long sqCounter = 0;                                // Couner that increments in intterup routine
unsigned long baseTimeS;                                    // Base time in sec from RTC to withc we sync the ms count

// RFID COMPONENT Create MFRC522 instance
MFRC522 mfrc522(RFID_CS, LCD_RST);                          // PH pins from PCB header file 
MFRC522::MIFARE_Key key;

// REAL TIME CLOCK COMPONENT
String timestamp, fulldate, fulltime; // yyyy-mm-dd and hh-mm-ss
unsigned long epoch;
//DateTime now;
time_t t;

// LCD COMPONENT
// LCD5110 myGLCD(4,5,6,A1,7); V2.0 green pcb
// LCD5110 myGLCD(4,5,6,7,8); (v1.0 blue PCB)
LCD5110 myGLCD(LCD_CLK, LCD_DIN , LCD_COM , LCD_RST, LCD_CE);   // PH pins from PCB header file
extern unsigned char TinyFont[];      // for 5110 LCD display
extern unsigned char SmallFont[];     // for 5110 LCD display
extern unsigned char MediumNumbers[]; // for 5110 LCD display
extern unsigned char BigNumbers[];    // for 5110 LCD display

// OTHER DEFINITIONS

long row = 3;               // Excel row counter
char value[16];             // written or read value
rExcel myExcel;             // class for Excel data exchanging
int led = 3;                        // define pin for LED on succesful write
int led1 = A0;                      // define pin for backup LED on succesful write
int LCDBL = A1;                     // define pin for LCD backlight 


// variables used for converting epoch time to readable time
unsigned long temp0=0, temp1=0, temp2=0, temp3=0, temp4=0,
              temp5=0, temp6=0, temp7=0, temp8=0, temp9=0,
              temp10=0, hours, mins, secs, MilliS;

              

//epoch time storage variables
unsigned long ss1Start=0, ss1Finish=0, SS1Time=0,   
              ss2Start=0, ss2Finish=0, SS2Time=0,
              ss3Start=0, ss3Finish=0, SS3Time=0,
              ss4Start=0, ss4Finish=0, SS4Time=0,
              ss5Start=0, ss5Finish=0, SS5Time=0,
              ss6Start=0, ss6Finish=0, SS6Time=0,
              SS1TimeMilliS=0, SS2TimeMilliS=0, SS3TimeMilliS=0,
              SS4TimeMilliS=0, SS5TimeMilliS=0, SS6TimeMilliS=0,
              totalRaceTime=0;

byte buffer[18];

void setup() {

  Wire.begin();
  Serial.begin(57600);                                  // Initialize serial communications with the PC
  myExcel.clearInput();                                 // rx buffer clearing
  myGLCD.InitLCD();                                     // initialise LCD
  setSyncProvider(RTC.get);                             // gets RTC time
  setSyncInterval(60);                                  // Set the number of seconds between re-syncs
  while (!Serial);                                      // Do nothing if no serial port is opened (added for Arduinos based on ATMEGA32U4)
  SPI.begin();                                          // Init SPI bus
  mfrc522.PCD_Init();                                   // Init MFRC522 card
  mfrc522.PCD_SetAntennaGain(mfrc522.RxGain_max);       // Enhance the MFRC522 Receiver Gain to maximum value of some 48 dB (default is 33dB)
      pinMode(LCDBL, OUTPUT); 
      pinMode(led, OUTPUT);
      pinMode(led1, OUTPUT);
        digitalWrite(LCDBL, LOW);                       // turn  LCD blacklight off
        digitalWrite(led, HIGH);                        // turn  LED off
        digitalWrite(led1, HIGH);                       // turn  backup LED off
  
  Serial.println(F("*** READY TO DOWNLOAD RESULTS DATA ***"));    //shows in serial that it is ready to read

  for (byte i = 0; i < 6; i++)                          // Prepare the key (used both as key A and as key B) using FFFFFFFFFFFFh which is the default at chip delivery from the factory
    { key.keyByte[i] = 0xFF; }
}
void loop() 
{
//      lcdScreen();                            // invoke LCD Screen fucntion
//      batteryVoltage();                       // invoke batteryVoltage Screen fucntion            
 
  if ( ! mfrc522.PICC_IsNewCardPresent())     // Look for new cards
      return;
 
  if ( ! mfrc522.PICC_ReadCardSerial())       // Select one of the cards
      return;

      readCardData();                         // Card is present and readble: invoke readCardData function
  
  mfrc522.PICC_HaltA();                       // Halt PICC
  mfrc522.PCD_StopCrypto1();                  // Stop encryption on PCD
}

void readCardData()
{

  byte size = sizeof(buffer);
  byte status;
  byte block;

  String string0;
  String string1;
  String string2;
  String string3;
  String string4;
  String string5;
  String string6;
  String string7;
  String string8;
  String string9;
  String string10;
  String string11;
  String string12;
  String string13;
  String string14;
  String string15;
  String fullstring;

/*
 * FIRST NAME // reads first name from PICC block 1, converts to characters and prints to Excel cells
 */

  block = 1;    

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

  string0 = char(buffer[0]);
  string1 = char(buffer[1]);
  string2 = char(buffer[2]);
  string3 = char(buffer[3]);
  string4 = char(buffer[4]);
  string5 = char(buffer[5]);
  string6 = char(buffer[6]);
  string7 = char(buffer[7]);
  string8 = char(buffer[8]);
  string9 = char(buffer[9]);
  string10 = char(buffer[10]);
  string11 = char(buffer[11]);
  string12 = char(buffer[12]);
  string13 = char(buffer[13]);
  string14 = char(buffer[14]);
  string15 = char(buffer[15]);

  fullstring = (string0 + string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15);

  char myFname[17];                                                // create an array of char
  fullstring.toCharArray(myFname,17);                              // copy fullstring content into myString
  myExcel.writeIndexed("DownloadResults", row,2, myFname);         // write firstname to Excel row, column 2
  myExcel.writeIndexed("DownloadResults", 1,43, myFname);          // write firstname to Excel row 1, column 43 (For printer Cell AQ1)

/*
 * SURNAME // reads surname from PICC block 2, converts to characters and prints to Excel cells
 */

block = 2;

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

  string0 = char(buffer[0]);
  string1 = char(buffer[1]);
  string2 = char(buffer[2]);
  string3 = char(buffer[3]);
  string4 = char(buffer[4]);
  string5 = char(buffer[5]);
  string6 = char(buffer[6]);
  string7 = char(buffer[7]);
  string8 = char(buffer[8]);
  string9 = char(buffer[9]);
  string10 = char(buffer[10]);
  string11 = char(buffer[11]);
  string12 = char(buffer[12]);
  string13 = char(buffer[13]);
  string14 = char(buffer[14]);
  string15 = char(buffer[15]);

  fullstring = (string0 + string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15);

  char mySname[17];  
  fullstring.toCharArray(mySname,17);                               // copy fullstring content into myString
  myExcel.writeIndexed("DownloadResults", row,3, mySname);          // write surname to Excel worksheet "DownloadResults" row, column 3
  myExcel.writeIndexed("DownloadResults", 2,43, mySname);           // write surname to Excel row 2, column 43 (For printer Cell AQ2)

/*
 * WATCH NUMBER //reads watch number from PICC block 4, converts to characters and prints to Excel cells
 */

block = 4;

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

  string0 = char(buffer[0]);
  string1 = char(buffer[1]);
  string2 = char(buffer[2]);
  string3 = char(buffer[3]);
  string4 = char(buffer[4]);
  string5 = char(buffer[5]);
  string6 = char(buffer[6]);
  string7 = char(buffer[7]);
  string8 = char(buffer[8]);
  string9 = char(buffer[9]);
  string10 = char(buffer[10]);
  string11 = char(buffer[11]);
  string12 = char(buffer[12]);
  string13 = char(buffer[13]);
  string14 = char(buffer[14]);
  string15 = char(buffer[15]);

    fullstring = (string0 + string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15);

char myRaceno[17];  
  fullstring.toCharArray(myRaceno,17);                                    // copy fullstring content into myString
  myExcel.writeIndexed("DownloadResults", row,1, myRaceno);               // write race number to Excel row, column 1
  myExcel.writeIndexed("DownloadResults", 3,43, myRaceno);                // write surname to Excel row 3, column 43 (For printer Cell AQ3)

/*
 * CATEGORY //Taken out as is unneeded becasue Excel takes care of the other fields but this can be put in if needed (or any other fields for that matter)
 */
/* 
block = 5;
// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

  string0 = char(buffer[0]);
  string1 = char(buffer[1]);
  string2 = char(buffer[2]);
  string3 = char(buffer[3]);
  string4 = char(buffer[4]);
  string5 = char(buffer[5]);
  string6 = char(buffer[6]);
  string7 = char(buffer[7]);
  string8 = char(buffer[8]);
  string9 = char(buffer[9]);
  string10 = char(buffer[10]);
  string11 = char(buffer[11]);
  string12 = char(buffer[12]);
  string13 = char(buffer[13]);
  string14 = char(buffer[14]);
  string15 = char(buffer[15]);

  fullstring = (string0 + string1 + string2 + string3 + string4 + string5 + string6 + string7 + string8 + string9 + string10 + string11 + string12 + string13 + string14 + string15);

char myCat[17]; 
  fullstring.toCharArray(myCat,17);                                     // copy fullstring content into myString
  myExcel.writeIndexed("DownloadResults", 4,43, myCat);                 // write category to Excel column 4
*/

/*
 * STAGE 1 START TIME  // reads SS1 Start Time from PICC  block 8 and stores in buffer
 */
 
block = 8;  

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store( still need to write to excel cell)


  buffer2epoch();
  ss1Start = temp10;


/*
 * STAGE 1 FINISH TIME  
 * // reads SS1 Finish Time from PICC  block 9, then subtracts ss1 finish from ss1 start and calculates SS1 Time
 * // and then downloads SS1 Time to Excel cells
 */
 
  block = 9;  

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }
          
// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss1Finish = temp10;

  writeSS1rawData();

  if (ss1Start > ss1Finish && ss1Finish !=0)                                        // if ss1Start bigger than ss1Finish and SS1Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 6, "ERROR");    // write ERROR to Excel, row, column 6
      myExcel.writeIndexed("DownloadResults", 5,43, "ERROR");      // write ERROR to Excel for Printer (AQ 5)
      SS1TimeMilliS = 0;                       // set stage time to zero for total count
    }
  else  
        if (ss1Start != 0 && ss1Finish != 0)
          { 
            SS1TimeMilliS = ss1Finish - ss1Start;
            MilliS=SS1TimeMilliS %1000;
            SS1Time = (SS1TimeMilliS - MilliS) / 1000;
            hours = (SS1Time/60/60);
            mins = (SS1Time-(hours*60*60))/60;
            secs = SS1Time-(hours*60*60)-(mins*60);
            
      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS  < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);    

     
      
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 6, myString);              // write myString (SS1 Time) to Excel row, column 6
          myExcel.writeIndexed("DownloadResults", 5,43, myString);                // write myString (SS1 Time) to Excel for Printer (AQ 5)

  
          }
              else
                {
                  myExcel.writeIndexed("DownloadResults", row, 6, "DNS/DNF");           // write DNS/DNF to Excel, row, column 6
                  myExcel.writeIndexed("DownloadResults", 5,43, "DNS/DNF");             // write DNS/DNF to Excel for Printer (AQ 5)
                  SS1TimeMilliS = 0;                                                          // set stage time to zero for total count
                }


/*
 * STAGE 2 START TIME // reads SS2 Start Time from PICC  block 12 and stores in buffer
 */

block = 12;
// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss2Start = temp10;
  
/*
 * STAGE 2 FINISH TIME 
 * // reads SS2 Finish Time from PICC  block 13, then subtracts ss2 finish from ss2 start and calculates SS2 Time
 * // and then downloads SS2 Time to Excel cells
 */
 
  block = 13;

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss2Finish = temp10;
  
  writeSS2rawData();
  
  if (ss2Start > ss2Finish && ss2Finish !=0)                                        // if ss2Start bigger than ss2Finish and SS2Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 7, "ERROR");    // write ERROR to Excel, row, column 7
      myExcel.writeIndexed("DownloadResults", 6,43, "ERROR");      // write ERROR to Excel for Printer (AQ 6)
      SS2TimeMilliS = 0;                                                                 // set stage time to zero for total count
    }
  else  
        if (ss2Start != 0 && ss2Finish != 0)
          { 
            SS2TimeMilliS = ss2Finish - ss2Start;
            MilliS=SS2TimeMilliS % 1000;
            SS2Time = (SS2TimeMilliS - MilliS) / 1000;
            hours = (SS2Time/60/60);
            mins = (SS2Time-(hours*60*60))/60;
            secs = SS2Time-(hours*60*60)-(mins*60);

      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS  < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);    
                
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                      // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 7, myString);                // write myString (SS2 Time) to Excel row, column 7
          myExcel.writeIndexed("DownloadResults", 6,43, myString);                  // write myString (SS2 Time) to Excel for Printer (AQ 6)

          }
        else
          {
            myExcel.writeIndexed("DownloadResults", row, 7, "DNS/DNF");             // write DNS/DNF to Excel, row, column 7
            myExcel.writeIndexed("DownloadResults", 6,43, "DNS/DNF");               // write DNS/DNF to Excel for Printer (AQ 6)
            SS2TimeMilliS = 0;                                                            // set stage time to zero for total count
          }




/*
 * STAGE 3 START TIME  // reads SS3 Start Time from PICC  block 16 and stores in buffer
*/

  block = 16;  // Stage 3 Start time

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss3Start = temp10;
  
/*
 * STAGE 3 FINISH TIME
 * // reads SS3 Finish Time from PICC  block 17, then subtracts ss1 finish from ss3 start and calculates SS3 Time
 * // and then downloads SS3 Time to Excel cells
 */
 
  block = 17;  // Stage 3 Finish time

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss3Finish = temp10;
  
  writeSS3rawData();

  if (ss3Start > ss3Finish && ss3Finish !=0)                                        // if ss3Start bigger than ss3Finish and SS3Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 8, "ERROR");    // write ERROR to Excel, row, column 8
      myExcel.writeIndexed("DownloadResults", 7,43, "ERROR");      // write ERROR to Excel for Printer (AQ 7)
      SS3TimeMilliS = 0;                                                                 // set stage time to zero for total count
    }
  else  

        if (ss3Start != 0 && ss3Finish != 0)
          {   
            SS3TimeMilliS = ss3Finish - ss3Start;
            MilliS=SS3TimeMilliS %1000;
            SS3Time = (SS3TimeMilliS - MilliS) / 1000;
            hours = (SS3Time/60/60);
            mins = (SS3Time-(hours*60*60))/60;
            secs = SS3Time-(hours*60*60)-(mins*60);
            
      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS  < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);    
          
      
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 8, myString);              // write myString (SS3 Time) to Excel row, column 8
          myExcel.writeIndexed("DownloadResults", 7,43, myString);                // write myString (SS3 Time) to Excel for Printer (AQ 7)


          }
        else
          {
            myExcel.writeIndexed("DownloadResults", row, 8, "DNS/DNF");           // write DNS/DNF to Excel, row, column 8
            myExcel.writeIndexed("DownloadResults", 7,43, "DNS/DNF");             // write DNS/DNF to Excel for Printer (AQ 7)
            SS3TimeMilliS = 0;                                                          // set stage time to zero for total count
          }


//  Serial.print("\t") ;             //tab separator

/*
 * STAGE 4 START TIME  // reads SS4 Start Time from PICC  block 20 and stores in buffer
 */
 
  block = 20;  // Stage 4 Start time
 
// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss4Start = temp10;
  
/*
 * STAGE 4 FINISH TIME
 * // reads SS4 Finish Time from PICC  block 21, then subtracts ss4 finish from ss4 start and calculates SS4 Time
 * // and then downloads SS4 Time to Excel cells
 */
 
  block = 21;   // Stage 4 Finish time

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss4Finish = temp10;
  
  writeSS4rawData();

  if (ss4Start > ss4Finish && ss4Finish !=0)                                        // if ss4Start bigger than ss4Finish and SS4Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 9, "ERROR");    // write ERROR to Excel, row, column 9
      myExcel.writeIndexed("DownloadResults", 8,43, "ERROR");      // write ERROR to Excel for Printer (AQ 8)
      SS4TimeMilliS = 0;                                                                 // set stage time to zero for total count
    }
  else  

        if (ss4Start != 0 && ss4Finish != 0)
          {   
            SS4TimeMilliS = ss4Finish - ss4Start;
            MilliS=SS4TimeMilliS %1000;
            SS4Time = (SS4TimeMilliS - MilliS) / 1000;
            hours = (SS4Time/60/60);
            mins = (SS4Time-(hours*60*60))/60;
            secs = SS4Time-(hours*60*60)-(mins*60);
      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS  < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);    
                
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 9, myString);              // write myString (SS4 Time) to Excel row, column 9
          myExcel.writeIndexed("DownloadResults", 8,43, myString);                // write myString (SS4 Time) to Excel for Printer (AQ 8)

          }
        else
          {
            myExcel.writeIndexed("DownloadResults", row, 9, "DNS/DNF");           // write DNS/DNF to Excel, row, column 9
            myExcel.writeIndexed("DownloadResults", 8,43, "DNS/DNF");             // write DNS/DNF to Excel for Printer (AQ 8)
            SS4TimeMilliS = 0;                                                          // set stage time to zero for total count
          }

//  Serial.print("\t") ;             //tab separator
  
/*
 * STAGE 5 START TIME  // reads SS5 Start Time from PICC  block 24 and stores in buffer
 */
 
  block = 24;  // Stage 5 Start time

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss5Start = temp10;
  
/*
 * STAGE 5 FINISH TIME
 * // reads SS5 Finish Time from PICC  block 25, then subtracts ss5 finish from ss5 start and calculates SS5 Time
 * // and then downloads SS5 Time to Excel cells
 */
 
  block = 25;     // Stage 5 Finish time

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss5Finish = temp10;
  
  writeSS5rawData();

    if (ss5Start > ss5Finish && ss5Finish !=0)                                        // if ss5Start bigger than ss5Finish and SS5Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 10, "ERROR");    // write ERROR to Excel, row, column 10
      myExcel.writeIndexed("DownloadResults", 9,43, "ERROR");      // write ERROR to Excel for Printer (AQ 9)
      SS5TimeMilliS = 0;                                                                 // set stage time to zero for total count
    }
  else  

        if (ss5Start != 0 && ss5Finish != 0)
          {   
            SS5TimeMilliS = ss5Finish - ss5Start;
            MilliS=SS5TimeMilliS %1000;
            SS5Time = (SS5TimeMilliS - MilliS) / 1000;
            hours = (SS5Time/60/60);
            mins = (SS5Time-(hours*60*60))/60;
            secs = SS5Time-(hours*60*60)-(mins*60);
      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS  < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);       
                
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 10, myString);             // write myString (SS5 Time) to Excel row, column 10
          myExcel.writeIndexed("DownloadResults", 9,43, myString);                // write myString (SS5 Time) to Excel for Printer (AQ 9)

          }
        else
          {
            myExcel.writeIndexed("DownloadResults", row, 10, "DNS/DNF");          // write DNS/DNF to Excel, row, column 10
            myExcel.writeIndexed("DownloadResults", 9,43, "DNS/DNF");             // write DNS/DNF to Excel for Printer (AQ 9)
            SS5TimeMilliS = 0;                                                          // set stage time to zero for total count
          }



/*
 * STAGE 6 START TIME  // reads SS6 Start Time from PICC  block 88 and stores in buffer
 */
 
  block = 28;

// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss6Start = temp10;
  
/*
 * STAGE 6 FINISH TIME
 * // reads SS6 Finish Time from PICC  block 29, then subtracts ss6 finish from ss6 start and calculates SS6 Time
 * // and then downloads SS6 Time to Excel cells
 */
 
  block = 29;

/// Authentication
  
  status = mfrc522.PCD_Authenticate(MFRC522::PICC_CMD_MF_AUTH_KEY_A,block, &key, &(mfrc522.uid));
    if (status != MFRC522::STATUS_OK) {
        Serial.print(F("PCD_Authenticate() failed: "));
        Serial.println(mfrc522.GetStatusCodeName(status));
        return; }
  
// Read data from the block into buffer

    status = mfrc522.MIFARE_Read(block, buffer, &size);
    if (status != MFRC522::STATUS_OK) 
      {   Serial.println(F("Download Failed"));
          return;     }

// Convert buffer type byte to unsigned long int and store

  buffer2epoch();
  ss6Finish = temp10;
  
  writeSS6rawData();

    if (ss6Start > ss6Finish && ss6Finish !=0)                                        // if ss6Start bigger than ss6Finish and SS6Finish is not equal to zero, then write 'error'
    {
      myExcel.writeIndexed("DownloadResults", row, 11, "ERROR");   // write ERROR to Excel, row, column 11
      myExcel.writeIndexed("DownloadResults", 10,43, "ERROR");     // write ERROR to Excel for Printer (AQ 10)
      SS6TimeMilliS = 0;                                                             // set stage time to zero for total count
    }
        if (ss6Start != 0 && ss6Finish != 0)
          { 
            SS6TimeMilliS = ss6Finish - ss6Start;
            MilliS = SS6TimeMilliS % 1000;
            SS6Time = (SS6TimeMilliS - MilliS) / 1000;
            hours = (SS6Time/60/60);
            mins = (SS6Time-(hours*60*60))/60;
            secs = SS6Time-(hours*60*60)-(mins*60);
            
      
          String timestamp;
      
          if (hours < 10) {timestamp += "0";}    
          timestamp += String(hours);
          timestamp += ":";
          if (mins < 10) {timestamp += "0";}    
          timestamp += String(mins);
          timestamp += ":";
          if (secs < 10) {timestamp += "0";}    
          timestamp += String(secs);
          timestamp += ".";
          if (MilliS < 10) {
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS > 9 && MilliS < 100) {
            timestamp += "0";
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);     

          
                
          char myString[17];    
          timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
          myExcel.writeIndexed("DownloadResults", row, 11, myString);             // write myString (SS6 Time) to Excel row, column 11
          myExcel.writeIndexed("DownloadResults", 10,43, myString);               // write myString (SS6 Time) to Excel for Printer (AQ 10)

          }
        else
          {
            myExcel.writeIndexed("DownloadResults", row, 11, "DNS/DNF");          // write DNS/DNF to Excel, row, column 11
            myExcel.writeIndexed("DownloadResults", 10,43, "DNS/DNF");            // write DNS/DNF to Excel for Printer (AQ 10)
            SS6TimeMilliS = 0;                                                          // set stage time to zero for total count
          }


/*
 * TOTAL TIME
 * Add all stage times and print the total to Excel
ToTal Time needs to be in MilliS eg: SS3TimeMilliS+SS4TimeMilliS etc
Rather write all this into a function to call upon and then convert 2 excel characters at the end
Look st Stage 3 Time calcs as an example
 */
/*
totalRaceTime = SS1TimeMilliS + SS2TimeMilliS + SS3TimeMilliS + SS4TimeMilliS + SS5TimeMilliS + SS6TimeMilliS;
hours = (totalRaceTime/(1000*60*60));
mins = (totalRaceTime-(hours*60*60))/60;
secs = totalRaceTime-(hours*60*60)-(mins*60);

MilliS = totalRaceTime %1000;*/
//    String testTRT , testmillis , testTRTS;
    char myString[17];   
            totalRaceTime = SS1TimeMilliS + SS2TimeMilliS + SS3TimeMilliS + SS4TimeMilliS + SS5TimeMilliS + SS6TimeMilliS;

//           testTRT += String(totalRaceTime);
//             testTRT.toCharArray(myString, 17); 
//            myExcel.writeIndexed("DownloadResults", row, 45, myString);             // write myString (totalRaceTime) to Excel row, column 12

            MilliS = totalRaceTime %1000;                                             //storing remainder of totalracetime in milliseconds
            
//             testmillis += String(MilliS);
//             testmillis.toCharArray(myString, 17); 
//            myExcel.writeIndexed("DownloadResults", row, 46, myString);             // write myString (totalRaceTime) to Excel row, column 12

            totalRaceTime=(totalRaceTime-MilliS)/1000;                               //converting totalracetime to whole seconds

//            testTRTS += String(totalRaceTime);
//            testTRTS.toCharArray(myString, 17); 
//            myExcel.writeIndexed("DownloadResults", row, 47, myString);             // write myString (totalRaceTime) to Excel row, column 12

            hours = (totalRaceTime/(60*60));
            mins = (totalRaceTime-(hours*60*60))/60;
            secs = totalRaceTime-(hours*60*60)-(mins*60);
            
    String timestamp;

    if (hours < 10) {timestamp += "0";}    
    timestamp += String(hours);
    timestamp += ":";
    if (mins < 10) {timestamp += "0";}    
    timestamp += String(mins);
    timestamp += ":";
    if (secs < 10) {timestamp += "0";}    
    timestamp += String(secs);
    timestamp += ".";
          if (MilliS< 10) {            //adding leading 00 for milliseconds
            timestamp += "00";
            timestamp += String(MilliS);
          }    
               
          else if (MilliS >9 && MilliS < 100) {
            timestamp += "0";                     //adding leading 0 for milliseconds
            timestamp += String(MilliS);
          }
          else timestamp += String(MilliS);       // no need to add leading zeros 
  
 
    timestamp.toCharArray(myString, 17);                                    // copy fullstring content into myString
    myExcel.writeIndexed("DownloadResults", row, 12, myString);             // write myString (totalRaceTime) to Excel row, column 12
    myExcel.writeIndexed("DownloadResults", 12,43, myString);               // write myString (totalRaceTime) to Excel for Printer (AQ 12)
    myExcel.save();                                                         // this doesn't save the excel file! Changed code (in excel) so that when this command is run, it actually copies the row to another sheet in excel (Dan)
  row++;
    Serial.println("*** SUCCESSFUL CARD READ ***");    
}

/* 
 *  Converts type byte to unsigned long int.
 *  Reads each digit from the buffer to the temp placeholders
 */
void buffer2epoch() {
//  byte size = sizeof(buffer);       // byte size = number of digits in the buffer
      temp0 = buffer[0]-'0';          //writes 1st digit of epoch time from buffer to temp0
      temp1 = buffer[1]-'0';          //writes 2nd digit of epoch time from buffer to temp1
      temp2 = buffer[2]-'0';          //writes 3rd digit of epoch time from buffer to temp2
      temp3 = buffer[3]-'0';          //writes 4th digit of epoch time from buffer to temp3
      temp4 = buffer[4]-'0';          //writes 5th digit of epoch time from buffer to temp4
      temp5 = buffer[5]-'0';          //writes 6th digit of epoch time from buffer to temp5
      temp6 = buffer[6]-'0';          //writes 7th digit of epoch time from buffer to temp6
      temp7 = buffer[7]-'0';          //writes 8th digit of epoch time from buffer to temp7
      temp8 = buffer[8]-'0';          //writes 9th digit of epoch time from buffer to temp8
      temp9 = buffer[9]-'0';          //writes 10th digit of epoch time from buffer to temp9
// use this for times with 7 digits. ie times written to watches between 00h00 and 02h42 include only hours,minutes, seconds and millis (initAfricaTimer()function in Timing POD code baseTimeS definition)
//        temp10 = ((temp0*1000000) + (temp1*100000) + (temp2*10000) + (temp3*1000) + (temp4*100) + (temp5*10) + temp6);
// use this for times with 8 digits. ie times written to watches  include only hours,minutes, seconds and millis (initAfricaTimer()function in Timing POD code baseTimeS definition)
         temp10 = ((temp0*10000000) + (temp1*1000000) + (temp2*100000) + (temp3*10000) + (temp4*1000) + (temp5*100) + (temp6*10) + temp7);
// use this for times with 9 digits. ie times written to watches  include day,hours,minutes, seconds and millis (initAfricaTimer()function in Timing POD code baseTimeS definition)
//        temp10 = ((temp0*100000000) + (temp1*10000000) + (temp2*1000000) + (temp3*100000) + (temp4*10000) + (temp5*1000) + (temp6*100) + (temp7*10) + temp8);

}
/*
void batteryVoltage()
{
          int sensorValue = analogRead(A2);                 // read the input on analog pin A2 for battery voltage reading:
          float voltage = (float)sensorValue / 72;          // Convert the analog reading to a voltage adjusted to Pieters multimeter
            if (voltage > 3.8)                              // If Voltage higher than 3.8 display "Good" on LCD
              { 
              myGLCD.setFont(SmallFont);
              myGLCD.print("Batt. Good",CENTER,40);
              }
              else                                          // If Voltage lower than 3.8 display "Turn off POD" on LCD and flash lights. POD will work until 3.5V but will stop working after that.
                {
                myGLCD.setFont(SmallFont);
                myGLCD.print("Turn Off POD",CENTER,40);     // display message and blink warning lights every half second (500ms)
                digitalWrite(led1, LOW); // turn  LED on
                delay(500);
                digitalWrite(led1, HIGH); // turn  LED off
                }
                    myGLCD.setFont(SmallFont);              // Displays batteryVoltage Info
                    myGLCD.print("Batt=",0,30);
                    myGLCD.printNumF(voltage, 1, CENTER,30);
                    myGLCD.print("volts",53,30);   

}

void lcdScreen()    // Displays main Screen Time & Pod station ID
{
    myGLCD.setFont(SmallFont);
    myGLCD.print(("DL 2 Excel"),CENTER,0);  // Displays POD ID 
    myGLCD.setFont(MediumNumbers);          // Displays Clock in Medium Numbers
    myGLCD.printNumI(hour(),0,10,2,'0');    // Minimum number of characters is 2 with on '0' to full the space if minimum is not reached
    myGLCD.printNumI(minute(),29,10,2,'0'); 
    myGLCD.printNumI(second(),58,10,2,'0'); 
    
    myGLCD.setFont(MediumNumbers);          // Displays colons
    myGLCD.drawRect(55,15,57,17);
    myGLCD.drawRect(55,20,57,22);
    myGLCD.drawRect(26,15,28,17);
    myGLCD.drawRect(26,20,28,22);
 
    myGLCD.update();
    myGLCD.clrScr();
            
}
*/
void writeSS1rawData()
{
          
          String ss1StartString;
          String ss1FinishString; 
          
          ss1StartString = String(ss1Start);           //converts ss1Start epoch time to string type and store in ss1StartString
          ss1FinishString = String(ss1Finish);         //converts ss1Finish epoch time to string type and store in ss1FinishString 
          
          char ss1ST[17];                               //declare character array for ss1 Start Time
          char ss1FT[17];                               //declare character array for ss1 Finish Time
          
          ss1StartString.toCharArray(ss1ST, 17);                           //convert ss1StartString to charater array ss1ST
          ss1FinishString.toCharArray(ss1FT, 17);                          //convert ss1FinishString to charater array ss1FT 
          
          myExcel.writeIndexed("DownloadResults", row, 28, ss1ST);           // write ss1Start time to Excel, row, column 28 
          myExcel.writeIndexed("DownloadResults", row, 29, ss1FT);           // write ss1Finish time to Excel, row, column 29
}

void writeSS2rawData()
{
          String ss2StartString;
          String ss2FinishString; 
          
          ss2StartString = String(ss2Start);           //converts ss2Start epoch time to string type and store in ss1StartString
          ss2FinishString = String(ss2Finish);         //converts ss2Finish epoch time to string type and store in ss1FinishString 
          
          char ss2ST[17];                               //declare character array for ss2 Start Time
          char ss2FT[17];                               //declare character array for ss2 Finish Time
          
          ss2StartString.toCharArray(ss2ST, 17);                           //convert ss2StartString to charater array ss2ST
          ss2FinishString.toCharArray(ss2FT, 17);                          //convert ss2FinishString to charater array ss2FT 
          
          myExcel.writeIndexed("DownloadResults", row, 30, ss2ST);           // write ss2Start time to Excel, row, column 30 
          myExcel.writeIndexed("DownloadResults", row, 31, ss2FT);           // write ss2Finish time to Excel, row, column 31
}

void writeSS3rawData()
{    
          String ss3StartString;
          String ss3FinishString; 
          
          ss3StartString += String(ss3Start);           //converts ss3Start epoch time to string type and store in ss3StartString
          ss3FinishString += String(ss3Finish);         //converts ss3Finish epoch time to string type and store in ss3FinishString 
          
          char ss3ST[17];                               //declare character array for ss3 Start Time
          char ss3FT[17];                               //declare character array for ss3 Finish Time
          
          ss3StartString.toCharArray(ss3ST, 17);                           //convert ss3StartString to charater array ss3ST
          ss3FinishString.toCharArray(ss3FT, 17);                          //convert ss3FinishString to charater array ss3FT 
          
          myExcel.writeIndexed("DownloadResults", row, 32, ss3ST);           // write ss3Start time to Excel, row, column 28 
          myExcel.writeIndexed("DownloadResults", row, 33, ss3FT);           // write ss3Finish time to Excel, row, column 29
}

void writeSS4rawData()
{
          String ss4StartString;
          String ss4FinishString; 
          
          ss4StartString += String(ss4Start);           //converts ss4Start epoch time to string type and store in ss1StartString
          ss4FinishString += String(ss4Finish);         //converts ss4Finish epoch time to string type and store in ss1FinishString 
          
          char ss4ST[17];                               //declare character array for ss4 Start Time
          char ss4FT[17];                               //declare character array for ss4 Finish Time
          
          ss4StartString.toCharArray(ss4ST, 17);                           //convert ss4StartString to charater array ss4ST
          ss4FinishString.toCharArray(ss4FT, 17);                          //convert ss4FinishString to charater array ss4FT 
          
          myExcel.writeIndexed("DownloadResults", row, 34, ss4ST);           // write ss4Start time to Excel, row, column 30 
          myExcel.writeIndexed("DownloadResults", row, 35, ss4FT);           // write ss4Finish time to Excel, row, column 31
}

void writeSS5rawData()
{
          String ss5StartString;
          String ss5FinishString; 
          
          ss5StartString += String(ss5Start);           //converts ss5Start epoch time to string type and store in ss1StartString
          ss5FinishString += String(ss5Finish);         //converts ss5Finish epoch time to string type and store in ss1FinishString 
          
          char ss5ST[17];                               //declare character array for ss5 Start Time
          char ss5FT[17];                               //declare character array for ss5 Finish Time
          
          ss5StartString.toCharArray(ss5ST, 17);                           //convert ss5StartString to charater array ss5ST
          ss5FinishString.toCharArray(ss5FT, 17);                          //convert ss5FinishString to charater array ss5FT 
          
          myExcel.writeIndexed("DownloadResults", row, 36, ss5ST);           // write ss5Start time to Excel, row, column 30 
          myExcel.writeIndexed("DownloadResults", row, 37, ss5FT);           // write ss5Finish time to Excel, row, column 31
}

void writeSS6rawData()
{
          String ss6StartString;
          String ss6FinishString; 
          
          ss6StartString += String(ss6Start);           //converts ss6Start epoch time to string type and store in ss1StartString
          ss6FinishString += String(ss6Finish);         //converts ss6Finish epoch time to string type and store in ss1FinishString 
          
          char ss6ST[17];                               //declare character array for ss6 Start Time
          char ss6FT[17];                               //declare character array for ss6 Finish Time
          
          ss6StartString.toCharArray(ss6ST, 17);                           //convert ss6StartString to charater array ss6ST
          ss6FinishString.toCharArray(ss6FT, 17);                          //convert ss6FinishString to charater array ss6FT 
          
          myExcel.writeIndexed("DownloadResults", row, 38, ss6ST);           // write ss6Start time to Excel, row, column 38
          myExcel.writeIndexed("DownloadResults", row, 39, ss6FT);           // write ss6Finish time to Excel, row, column 39  
}

