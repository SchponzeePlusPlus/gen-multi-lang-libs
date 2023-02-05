/*
	Program: Solar Mobility System (Concept 0 Version 0) Panel Control Program (Revision 0) (.INO/C++)
	Module: Custom SMS Objects Header File (Header Only)
	Authors: Team SLK - Nicholas Kazanidis (101097401), Kayla Lai (100588247), Leonard Sponza (100588917)
	Created On: 17/08/2020
	Last Modified: 12/11/2020
	Branch: feat-comments-00

	Description:
	Contains the project source code responsible for communicating with the data logging expansion shield, which features an SD card and an RTC.
*/

#ifndef INCLUDES_DATALOGGINGMODULE_H_
#define INCLUDES_DATALOGGINGMODULE_H_

// Call necessary modules/libraries via header files here
// Call hardware specific libraries here
#include "Arduino.h"

// Call component specific program modules
#include "ldrmodule.h"
// More relevant modules are called in .cpp file

// Specify a namespace here
// No namespaces are specified in this file

// Declare class objects here
// No objects are created in this file

// Declare enumerators here
// No enumerators are declared in this file

// Declare global constants here
// "Magic numbers" are assigned as global constants
//const String dataLogFileName = "SMS_DataLog.csv";
// const signed long long int MAX_MINUTE_INTERVAl_PER_DL_REC = 60;
// Declare a global constant to specify the scheduling interval of the data logging routine
const signed long long int MAX_MINUTE_INTERVAl_PER_DL_REC = 1;

//enum DataLoggingState

// const String YEAR_NAME = "Year";
// const String MONTH_NAME = "Month";
// const String DAY_NAME = "Day";
// const String HOUR_NAME = "Hour";
// const String MIN_NAME = "Minute";
// const String SEC_NAME = "Second";

const char YEAR_NAME[] = "Year";
const char MONTH_NAME[] = "Month";
const char DAY_NAME[] = "Day";
const char HOUR_NAME[] = "Hour";
const char MIN_NAME[] = "Minute";
const char SEC_NAME[] = "Second";

// const PROGMEM char YEAR_NAME[] = "Year";
// const PROGMEM char MONTH_NAME[] = "Month";
// const PROGMEM char DAY_NAME[] = "Day";
// const PROGMEM char HOUR_NAME[] = "Hour";
// const PROGMEM char MIN_NAME[] = "Minute";
// const PROGMEM char SEC_NAME[] = "Second";

// Define records or structs
struct dateTimeRecord
{
	uint16_t Year;
	uint8_t Month;
	uint8_t Day;
	uint8_t Hour;
	uint8_t Min;
	uint8_t Sec;
};

// const String DATA_LOG_ID_NAME = "Data Log ID";
// const String MEASURED_ROT_ANGLE_DAILY_NAME = "Measured Daily Rotation Angle (deg)";
// const String MEASURED_ROT_ANGLE_SEASONAL_NAME = "Measured Seasonal Rotation Angle (deg)";
// const String CUR_SENSOR_VAL_SMS_NAME = "sms mA";
// const String POW_SENSOR_VAL_SMS_NAME = "SMS Power Sensor Value (mW)";
// const String CUR_SENSOR_VAL_STATIC_NAME = "static mA";
// const String POW_SENSOR_VAL_DUMMY_NAME = "Dummy Power Sensor Value (mW)";

const char DATA_LOG_ID_NAME[] = "Data Log ID";
const char MEASURED_ROT_ANGLE_DAILY_NAME[] = "Measured Daily Rotation Angle (deg)";
const char MEASURED_ROT_ANGLE_SEASONAL_NAME[] = "Measured Seasonal Rotation Angle (deg)";
const char CUR_SENSOR_VAL_SMS_NAME[] = "sms mA";
const char POW_SENSOR_VAL_SMS_NAME[] = "SMS Power Sensor Value (mW)";
const char CUR_SENSOR_VAL_STATIC_NAME[] = "static mA";
const char POW_SENSOR_VAL_DUMMY_NAME[] = "Dummy Power Sensor Value (mW)";

// const PROGMEM char DATA_LOG_ID_NAME[] = "Data Log ID";
// const PROGMEM char MEASURED_ROT_ANGLE_DAILY_NAME[] = "Measured Daily Rotation Angle (deg)";
// const PROGMEM char MEASURED_ROT_ANGLE_SEASONAL_NAME[] = "Measured Seasonal Rotation Angle (deg)";
// const PROGMEM char CUR_SENSOR_VAL_SMS_NAME[] = "SMS Current Sensor Value (mA)";
// const PROGMEM char POW_SENSOR_VAL_SMS_NAME[] = "SMS Power Sensor Value (mW)";
// const PROGMEM char CUR_SENSOR_VAL_STATIC_NAME[] = "Static Current Sensor Value (mA)";
// const PROGMEM char POW_SENSOR_VAL_DUMMY_NAME[] = "Static Power Sensor Value (mW)";

// Declare a struct that contains the buffer of the next data log to be appended to the CSV file on the SD card
struct dataLogRecord
{
	int dataLogID;
	struct dateTimeRecord dateTimeStamp;
	struct LDRRefRVoltRecord LDRVRec;
	float measuredRotationAngleDaily;
	float measuredRotationAngleSeasonal;
	float currentSensorValSMS;
	float powerSensorValSMS;
	float currentSensorValStatic;
	float powerSensorValDummy;
};

/**
	@brief Configure SD card via SPI protocol
*/
void configSD();

/**
	@brief Open a CSV file to write mode on the SD card
*/
void openSDCSVFile();

/**
	@brief Close the data log file
*/
void closeDataLogFile();

/**
	@brief Configure the Real-Time-Clock (RTC) via I2C protocol
*/
void configRTC();

/**
	@brief Read current data and time from RTC
	@return RTC date and timestamp
*/
struct dateTimeRecord readRTC();

/**
	@brief Nullifies a date and time record struct
    @return Null date and time record struct
*/
struct dateTimeRecord nullifyDateTimeRec();

/**
	@brief Prints a date and time record struct to Serial Monitor
	@param dateTimeRec Date and time record struct
*/
void printDateTime(struct dateTimeRecord dateTimeRec);

/**
	@brief Calculates the time difference between the current date and time and a previously stored date and time
	@param currentDateTime Current date and time record struct
	@param previousDateTime Previous date and time record struct
	@return Date and time difference
*/
struct dateTimeRecord calcTimeElapsed(struct dateTimeRecord currentDateTime, struct dateTimeRecord previousDateTime);

/**
	@brief Converts the date and time difference into minutes elapsed, does not consider seconds
	@param dateTimeElapsedRec Date and time difference
	@return Minutes elapsed between two date and timestamps
*/
signed long long int calcMinsElapsed(struct dateTimeRecord dateTimeElapsedRec);

/**
	@brief Prepares the next entry for data logging into a record
	@param dataLogCount Data log counter
	@param dateTime Date and time stamp
	@param LDRRefRVoltRec LDR reference resistor voltage record struct
	@param measuredRotAngleDaily Daily actuator measured rotation angle
	@param measuredRotAngleSeasonal Seasonal actuator measured rotation angle
	@param PVCSensorCurSMS PVC sensor current value for the solar panels on the SMS
	@param PVCSensorValSMS PVC sensor power value for the solar panels on the SMS
	@param PVCSensorCurStatic PVC sensor current value for the solar panels on the static benchmark
	@param PVCSensorValDummy PVC sensor power value for the solar panels on the static benchmark
	@return Data log record struct
*/
struct dataLogRecord prepareDataLogRecord(int dataLogCount, struct dateTimeRecord dateTime, struct LDRRefRVoltRecord LDRRefRVoltRec, float measuredRotAngleDaily, float measuredRotAngleSeasonal, float PVCSensorCurSMS, float PVCSensorValSMS, float PVCSensorCurStatic, float PVCSensorValDummy);

/**
	@brief Nullifies data log record struct
    @return Null data log record struct
*/
struct dataLogRecord nullifyDataLogRecord();

/**
	@brief Appends the data log column names to the CSV file on the SD card
*/
void writeDataLogColNames();

/**
	@brief Appends the data log record struct values to the CSV file on the SD card
	@param dataLogRec Data log record struct
*/
void writeDataLogRec(struct dataLogRecord dataLogRec);

/**
	@brief Checks whether data logging is due to occur
	@param minutesElapsed Minutes elapsed since the last data log
	@param dataLogCount Data log counter
	@return True if data logging is due
*/
bool checkDataLoggingInterval(signed long long int minutesElapsed, int dataLogCount);

/**
	@brief Checks whether a scheduled event is due to occur
	@param minutesElapsed Minutes elapsed since the last occurrence of the scheduled event
	@param schedulingInterval Specified interval of time to pass before the next occurrence of the scheduled event
	@param intervalCnt Event occurrence counter
	@param chkIntervFirstCount Check Interval of the first count
	@return True if scheduled event is due
*/
bool checkSchedulingInterval(signed long long int minutesElapsed, int schedulingInterval, int intervalCnt, bool chkIntervFirstCount);

#endif
