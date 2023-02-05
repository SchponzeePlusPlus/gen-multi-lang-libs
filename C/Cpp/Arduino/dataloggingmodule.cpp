/*
	Program: Solar Mobility System (Concept 0 Version 0) Panel Control Program (Revision 0) (.INO/C++)
	Module: Data Logging
	Authors: Team SLK - Nicholas Kazanidis (101097401), Kayla Lai (100588247), Leonard Sponza (100588917)
	Created On: 17/08/2020
	Last Modified: 29/09/2020
	Revision: 01
	Branch: SponzaPlusPlus-patch-05

	Description:
	XC4536
*/

// Call necessary modules/libraries via header files

#include "Arduino.h"

// SD Card
 #include <SPI.h>
 #include <SD.h>
 #include "RTClib.h"
 #include "Wire.h"

#include "customlibrary.h"
#include "ldrmodule.h"
#include "dataloggingmodule.h"
#include "dataloggingmodobjects.h"
//all these diodes red is the cathode

// Declare records or structs
/* struct structname
{
	string ;
	int ;
	float ;
}; */

// Declare global constants, "magic numbers" are assigned as global constants
//const  =;

void configSD()
{
//  #ifndef ESP8266
//    while (!Serial); // wait for serial port to connect. Needed for native USB
//  #endif
  
	Serial.print(F("Initializing SD card..."));
	pinMode(SS, OUTPUT);
	if (!SD.begin(10, 11, 12, 13))
	{
		Serial.println(F("Card failed, or not present"));
		// don't do anything more:
		while (1);
  	}
	Serial.println(F("card initialized."));

////  SDDataFile = SD.open("datalog2.csv", FILE_WRITE);
////  SDDataFile = SD.open(dataLogFileName, FILE_WRITE);
////  SDDataFile = SD.open("SMSDataLog.csv", FILE_WRITE);
////  File DataFile = SD.open("SMSDataLog.csv", FILE_WRITE);
//  SDDataFile = SD.open("datalog4.csv", FILE_WRITE);
//  Serial.println(testVal);
//  if (! SDDataFile)
////  if (!DataFile)
//  {
//    // Serial.println("error opening datalog2.csv");
//    Serial.print("error opening ");
//    Serial.println("CSV File");
////    Serial.println(dataLogFileName);
//    // Wait forever since we cant write data
//    while (1) ;
//  }
}

void openSDCSVFile()
{
//  SDDataFile = SD.open("datalog2.csv", FILE_WRITE);
//  SDDataFile = SD.open(dataLogFileName, FILE_WRITE);
//  SDDataFile = SD.open("SMSDataLog.csv", FILE_WRITE);
//	if (! SDDataFile)
//	{
//		// Serial.println("error opening datalog2.csv");
//		Serial.print("error opening ");
//    Serial.println("CSV File");
////		Serial.println(dataLogFileName);
//		// Wait forever since we cant write data
//		while (1) ;
//	}

  //  SDDataFile = SD.open("datalog2.csv", FILE_WRITE);
//  SDDataFile = SD.open(dataLogFileName, FILE_WRITE);
//  SDDataFile = SD.open("SMSDataLog.csv", FILE_WRITE);
//  File DataFile = SD.open("SMSDataLog.csv", FILE_WRITE);
    SDDataFile = SD.open("datalog4.csv", FILE_WRITE);
//    Serial.println(testVal);
    if (! SDDataFile)
    //  if (!DataFile)
    {
      // Serial.println("error opening datalog2.csv");
      Serial.print(F("error opening "));
      Serial.println(F("CSV File"));
    //    Serial.println(dataLogFileName);
      // Wait forever since we cant write data
      while (1) ;
    }
}

void closeDataLogFile()
{
	SDDataFile.close();
}

void configRTC()
{
  if (! rtc.begin())
  {
    Serial.println(F("Couldn't find RTC"));
    Serial.flush();
    abort();
  }
	if (! rtc.isrunning())
	{
		//Serial.println("RTC is NOT running, let's set the time!");
		Serial.println(F("RTC is NOT running!"));
		// When time needs to be set on a new device, or after a power loss, the
		// following line sets the RTC to the date & time this sketch was compiled
		//rtc.adjust(DateTime(F(__DATE__), F(__TIME__)));
		// This line sets the RTC with an explicit date & time, for example to set
		// January 21, 2014 at 3am you would call:
		// rtc.adjust(DateTime(2014, 1, 21, 3, 0, 0));
	}

	// When time needs to be re-set on a previously configured device, the
	// following line sets the RTC to the date & time this sketch was compiled
	Serial.println(F("Let's set the RTC's time to the timestamp this sketch was compiled at!"));
	rtc.adjust(DateTime(F(__DATE__), F(__TIME__)));
	// This line sets the RTC with an explicit date & time, for example to set
	// January 21, 2014 at 3am you would call:
	// rtc.adjust(DateTime(2014, 1, 21, 3, 0, 0));
}

struct dateTimeRecord readRTC()
{
	struct dateTimeRecord result;

//	DateTime rtcInput = rtc.now();
  rtcInput = rtc.now();
//   now = rtc.now();

	result.Year = rtcInput.year();
	result.Month = rtcInput.month();
  result.Day = rtcInput.day();
  result.Hour = rtcInput.hour();
  result.Min = rtcInput.minute();
  result.Sec = rtcInput.second();

//    result.Year = now.year();
//    Serial.println(result.Year);
//    result.Month = now.month();
//    result.Day = now.day();
//    result.Hour = now.hour();
//    result.Min = now.minute();
//    result.Sec = now.second();

	return result;
}

struct dateTimeRecord nullifyDateTimeRec()
{
	struct dateTimeRecord result;

	result.Year = 0;
	result.Month = 0;
	result.Day = 0;
	result.Hour = 0;
	result.Min = 0;
	result.Sec = 0;

	return result;
}

void printDateTime(struct dateTimeRecord dateTimeRec)
{
	Serial.println(F("Print date time rec"));
    Serial.print(dateTimeRec.Year, DEC);
    Serial.print('/');
    Serial.print(dateTimeRec.Month, DEC);
    Serial.print('/');
    Serial.print(dateTimeRec.Day, DEC);
    Serial.print(F(" ("));
    Serial.print(dateTimeRec.Hour, DEC);
    Serial.print(':');
    Serial.print(dateTimeRec.Min, DEC);
    Serial.print(':');
    Serial.print(dateTimeRec.Sec, DEC);
    Serial.print(F(") "));
    Serial.println();
}

struct dateTimeRecord calcTimeElapsed(struct dateTimeRecord currentDateTime, struct dateTimeRecord previousDateTime)
{
	struct dateTimeRecord result;

	result.Year = currentDateTime.Year - previousDateTime.Year;
	result.Month = currentDateTime.Month - previousDateTime.Month;
	result.Day = currentDateTime.Day - previousDateTime.Day;
	result.Hour = currentDateTime.Hour - previousDateTime.Hour;
	result.Min = currentDateTime.Min - previousDateTime.Min;
	result.Sec = currentDateTime.Sec - previousDateTime.Sec;

	return result;
}

signed long long int calcMinsElapsed(struct dateTimeRecord dateTimeElapsedRec)
{
	signed long long int result;

	result = (dateTimeElapsedRec.Min + (60 *
			(dateTimeElapsedRec.Hour + (24 *
			(dateTimeElapsedRec.Day + ((365 / 12) *
			(dateTimeElapsedRec.Month + (12 *
			(dateTimeElapsedRec.Year)))))))));

	return result;
}

struct dataLogRecord prepareDataLogRecord(int dataLogCount, struct dateTimeRecord dateTime, struct LDRRefRVoltRecord LDRRefRVoltRec, float measuredRotAngleDaily, float measuredRotAngleSeasonal, float PVCSensorCurSMS, float PVCSensorValSMS, float PVCSensorCurStatic, float PVCSensorValDummy)
{
	struct dataLogRecord result;

	result.dataLogID = dataLogCount;
	result.dateTimeStamp = dateTime;
	result.LDRVRec = LDRRefRVoltRec;
	// float LDRVoltageNE;
	// float LDRVoltageSE;
	// float LDRVoltageSW;
	// float LDRVoltageNW;
	result.measuredRotationAngleDaily = measuredRotAngleDaily;
	result.measuredRotationAngleSeasonal = measuredRotAngleSeasonal;
	result.currentSensorValSMS = PVCSensorCurSMS;
	result.powerSensorValSMS = PVCSensorValSMS;
	result.currentSensorValStatic = PVCSensorCurStatic;
	result.powerSensorValDummy = PVCSensorValDummy;

	return result;
}

struct dataLogRecord nullifyDataLogRecord()
{
	struct dataLogRecord result;

	result.dataLogID = 0;
	result.dateTimeStamp.Year = 0;
	result.dateTimeStamp.Month = 0;
	result.dateTimeStamp.Day = 0;
	result.dateTimeStamp.Hour = 0;
	result.dateTimeStamp.Min = 0;
	result.dateTimeStamp.Sec = 0;
	result.LDRVRec.NEVolt = 0.0;
	result.LDRVRec.SEVolt = 0.0;
	result.LDRVRec.SWVolt = 0.0;
	result.LDRVRec.NWVolt = 0.0;
	// float LDRVoltageNE;
	// float LDRVoltageSE;
	// float LDRVoltageSW;
	// float LDRVoltageNW;
	result.measuredRotationAngleDaily = 0.0;
	result.measuredRotationAngleSeasonal = 0.0;
	result.currentSensorValSMS = 0.0;
	result.powerSensorValSMS = 0.0;
	result.currentSensorValStatic = 0.0;
	result.powerSensorValDummy = 0.0;

	return result;
}

//void writeDataLogColNames(struct dateTimeColNames ColNamesRec)
void writeDataLogColNames()
{
//	String columnNames = DATA_LOG_ID_NAME + "," + YEAR_NAME + "," + DATE_TIME_COL_NAMES_REC.monthName  + "," + DATE_TIME_COL_NAMES_REC.dayName + "," + DATE_TIME_COL_NAMES_REC.hourName + "," + DATE_TIME_COL_NAMES_REC.minName + "," + DATE_TIME_COL_NAMES_REC.secName
//    + "," + NE_VOLT_COL_NAME + "," + SE_VOLT_COL_NAME + "," + SW_VOLT_COL_NAME + "," + NW_VOLT_COL_NAME + "," + MEASURED_ROT_ANGLE_DAILY_NAME + "," + MEASURED_ROT_ANGLE_SEASONAL_NAME + "," + POW_SENSOR_VAL_SMS_NAME + "," + POW_SENSOR_VAL_DUMMY_NAME;
//   String columnNames = DATA_LOG_ID_NAME + "," + YEAR_NAME + "," + MONTH_NAME  + "," + DAY_NAME + "," + HOUR_NAME + "," + MIN_NAME + "," + SEC_NAME
//     + "," + NE_VOLT_COL_NAME + "," + SE_VOLT_COL_NAME + "," + SW_VOLT_COL_NAME + "," + NW_VOLT_COL_NAME + "," + MEASURED_ROT_ANGLE_DAILY_NAME + "," + MEASURED_ROT_ANGLE_SEASONAL_NAME + "," + CUR_SENSOR_VAL_SMS_NAME + "," + POW_SENSOR_VAL_SMS_NAME + "," + CUR_SENSOR_VAL_STATIC_NAME + "," + POW_SENSOR_VAL_DUMMY_NAME;
	String columnNames = String(DATA_LOG_ID_NAME) + "," + String(YEAR_NAME) + "," + String(MONTH_NAME)  + "," + String(DAY_NAME) + "," + String(HOUR_NAME) + "," + String(MIN_NAME) + "," + String(SEC_NAME)
		+ "," + String(NE_VOLT_COL_NAME) + "," + String(SE_VOLT_COL_NAME) + "," + String(SW_VOLT_COL_NAME) + "," + String(NW_VOLT_COL_NAME) + "," + String(MEASURED_ROT_ANGLE_DAILY_NAME) + "," + String(MEASURED_ROT_ANGLE_SEASONAL_NAME) + "," + String(CUR_SENSOR_VAL_SMS_NAME) + "," + String(POW_SENSOR_VAL_SMS_NAME) + "," + String(CUR_SENSOR_VAL_STATIC_NAME) + "," + String(POW_SENSOR_VAL_DUMMY_NAME);
	
	SDDataFile.println(columnNames);

  Serial.println(F("Printing Column Names:"));
  Serial.println(columnNames);
}

void writeDataLogRec(struct dataLogRecord dataLogRec)
{
	String currentSDRecord;

	currentSDRecord =
	(
		String(dataLogRec.dataLogID) + "," + String(dataLogRec.dateTimeStamp.Year) + "," + String(dataLogRec.dateTimeStamp.Month) + "," + String(dataLogRec.dateTimeStamp.Day)
    	+ "," + String(dataLogRec.dateTimeStamp.Hour) + "," + String(dataLogRec.dateTimeStamp.Min) + "," + String(dataLogRec.dateTimeStamp.Sec) + "," + String(dataLogRec.LDRVRec.NEVolt)
    	+ "," + String(dataLogRec.LDRVRec.SEVolt) + "," + String(dataLogRec.LDRVRec.SWVolt) + "," + String(dataLogRec.LDRVRec.NWVolt) + "," + String(dataLogRec.measuredRotationAngleDaily)
    	+ "," + String(dataLogRec.measuredRotationAngleSeasonal) + "," + String(dataLogRec.currentSensorValSMS)+ "," + String(dataLogRec.powerSensorValSMS)
		+ "," + String(dataLogRec.currentSensorValStatic) + "," + String(dataLogRec.powerSensorValDummy)
	);
	
	SDDataFile.println(currentSDRecord);
  
    SDDataFile.flush();

	// Serial.println(F("Printing Record:"));
	Serial.println(("Printing Record:"));
  	Serial.println(currentSDRecord);
}

bool checkDataLoggingInterval(signed long long int minutesElapsed, int dataLogCount)
{
	bool result;
	if ((minutesElapsed >= MAX_MINUTE_INTERVAl_PER_DL_REC) || (dataLogCount == 0))
		{
			result = true;
			//previousDataLogTimestamp = currentData.timestamp;
			//minsElapsed = 0;
		}
	else
	{
		result = false;
	}

	return result;
}

bool checkSchedulingInterval(signed long long int minutesElapsed, int schedulingInterval, int intervalCnt, bool chkIntervFirstCount)
{
	bool result;
	if ((minutesElapsed >= schedulingInterval) || ((chkIntervFirstCount) && (intervalCnt == 0)))
		{
			result = true;
		}
	else
	{
		result = false;
	}

	return result;
}
