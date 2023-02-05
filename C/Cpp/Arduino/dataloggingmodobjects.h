/*
	Program: Solar Mobility System (Concept 0 Version 0) Panel Control Program (Revision 0) (.INO/C++)
	Module: Data Logging Objects Header File
	Authors: Team SLK - Nicholas Kazanidis (101097401), Kayla Lai (100588247), Leonard Sponza (100588917)
	Created On: 29/09/2020
	Last Modified: 29/09/2020
	Revision: 01
	Branch: SponzaPlusPlus-patch-05

	Description:
	Jaycar duinotech XC4536 Data Logging (Expansion) Shield, which features an SD Card Slot (interfaces almost directly to Arduino with 4 Digital pins) and a DS1307 Real-Time-Clock (RTC) IC with CR1220 battery backup.
*/

#ifndef INCLUDES_DATALOGGINGMODOBJECTS_H_
#define INCLUDES_DATALOGGINGMODOBJECTS_H_

#include "Arduino.h"

// SD Card
#include <SPI.h>
#include <SD.h>
#include "RTClib.h"
#include "Wire.h"

// RTC_DS1307 RTC_OBJ;
// DateTime TIME_OBJ;
// File SD_DATA_FILE_OBJ;
RTC_DS1307 rtc;
// DateTime dateTimeObj;
DateTime rtcInput;
//DateTime now;
File SDDataFile;

int testVal = 5;

//String dataLogFileName = "SMS_DataLog.csv";
//char dataLogFileName[] = "SMS_DataLog.csv"; 

//struct DATE_TIME_COL_NAMES_REC
//struct DateTimeColNamesRec
//{
//  const String yearName = "Year";
//  const String monthName = "Month";
//  const String dayName = "Day";
//  const String hourName = "Hour";
//  const String minName = "Minute";
//  const String secName = "Second";
////};
//} DATE_TIME_COL_NAMES_REC;

#endif
