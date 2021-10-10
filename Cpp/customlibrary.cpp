/*
	Program: Solar Mobility System (Concept 0 Version 0) Panel Control Program (Revision 0) (.INO/C++)
	Module: Custom Library
	Authors: Team SLK - Nicholas Kazanidis (101097401), Kayla Lai (100588247), Leonard Sponza (100588917)
	Created On: 17/08/2020
	Last Modified: 17/08/2020
	Branch: SponzaPlusPlus-patch-0

	Description:
	... 
*/

// Call necessary modules/libraries via header files
/* #include <algorithm>
#include <array>
// Access cmath library
#include <cmath>
#include <ctime>
#include <cstdlib>
// File stream
#include <fstream>
#include <iomanip>
// Access input output related code
#include <iostream>
#include <limits>
// for integer to character array operations
#include <sstream>
// Access string related code
#include <string>
#include <vector>

#include <ctype.h>
#include <math.h>
#include <stdio.h>
#include <stdlib.h>
#include <time.h> */

//#include "String.h"

#include "Arduino.h"
#include "customlibrary.h"

// Use the C++ standard namespace which includes cin and cout
/* using namespace std;
 */


// Declare records or structs
/* struct structname
{
	string ;
	int ;
	float ;
}; */

// Declare global constants, "magic numbers" are assigned as global constants
//const  =;

/* const char* convIntToCharArray(int inputInt)
{
	std::stringstream tempCharArray;
		tempCharArray << (inputInt);

			std::string charArrayString = tempCharArray.str();
			const char* outputCharArray = charArrayString.c_str();
	return outputCharArray;
} */

// float scaleAdcFloatInput(unsigned rawAdcInput, unsigned AdcInputBitRes, float outputValueMin, float outputValueMax)
float scaleAdcFloatInput(int rawAdcInput, int AdcInputBitResMax, int AdcInputBitResMin, float outputValueMax, float outputValueMin)
{
	float outputValueRange = outputValueMax - outputValueMin;
	int ADCInputBitResRange = AdcInputBitResMax - AdcInputBitResMin;
	// float result = ((rawAdcInput / AdcInputBitResMax) * outputValueRange) / 100;
	//float result = ((((float) rawAdcInput) / ((float) AdcInputBitResMax)) * outputValueRange);
	float result = ((((float) rawAdcInput) / ((float) ADCInputBitResRange)) * outputValueRange);
	return result;
}

// gradient finders

//inputValueType printDebugValue(bool debugMode, String valueName, inputValueType inputValue, String valueUnit)
//{
//	if(debugMode)
//	{
//		Serial.print(valueName);
//		Serial.print(" value: ")
//		Serial.print(inputValue);
//		Serial.print(" ");
//		Serial.println(valueUnit);
//	}
//	else
//	{
//		// nothing
//	}
//	return 0;
//}

//void printFloatDebugValue(bool debugMode, String valueName, float inputFloatValue, String valueUnit)
//{
//	if(debugMode)
//	{
//		Serial.print(valueName);
//		Serial.print(" value: ");
//		Serial.print(inputFloatValue);
//		Serial.print(" ");
//		Serial.println(valueUnit);
//	}
//	else
//	{
//		// nothing
//	}
//}
//
//void printIntDebugValue(bool debugMode, String valueName, int inputIntValue, String valueUnit)
//{
//	if(debugMode)
//	{
//		Serial.print(valueName);
//		Serial.print(" value: ");
//		Serial.print(inputIntValue);
//		Serial.print(" ");
//		Serial.println(valueUnit);
//	}
//	else
//	{
//		// nothing
//	}
//}
//
//void printBoolDebugValue(bool debugMode, String valueName, bool inputBoolValue)
//{
//	if(debugMode)
//	{
//		Serial.print(valueName);
//		Serial.print(" value: ");
//		if(inputBoolValue)
//		{
//			Serial.println("TRUE");
//		}
//		else
//		{
//			Serial.println("FALSE");
//		}
//	}
//	else
//	{
//		// nothing
//	}
//}
