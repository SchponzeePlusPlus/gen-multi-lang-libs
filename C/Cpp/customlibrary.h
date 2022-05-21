/*
	Program: Solar Mobility System (Concept 0 Version 0) Panel Control Program (Revision 0) (.INO/C++)
	Module: Custom Library Header File
	Authors: Team SLK - Nicholas Kazanidis (101097401), Kayla Lai (100588247), Leonard Sponza (100588917)
	Created On: 17/08/2020
	Last Modified: 12/11/2020
	Branch: feat-comments-00

	Description:
	... 
*/

#ifndef INCLUDES_CUSTOMLIBRARY_H_
#define INCLUDES_CUSTOMLIBRARY_H_

// Call necessary modules/libraries via header files here
// Call hardware specific libraries here
#include "Arduino.h"
//#include "String.h"

// Specify a namespace here
// No namespaces are specified in this file

// Declare a template here
//template <typename inputValueType>

// Declare enumerators here
// No enumerators are declared in this file

// Define records or structs here
// No structs are defined in this file

// Declare classes here?
// No classess are declared in this file?

// Declare global constants here
// "Magic numbers" are assigned as global constants

// Declare a character array for a commonly used string in troubleshooting, to save memory
const char COLON_SPACE_CHAR_ARR[] = ": ";
// const PROGMEM char COLON_SPACE_CHAR_ARR[] = ": ";

// const PROGMEM char BITS_UNIT_NAME[] = "bits";

const PROGMEM char ADC_RES_MAX_NAME[] = "ADC Input Resolution Max Value (bits)";
const PROGMEM char ADC_RES_MIN_NAME[] = "ADC Input Resolution Min Value (bits)";
const PROGMEM char ADC_OUT_RANGE_NAME[] = "ADC Output Value Range";
const PROGMEM char ADC_IN_PORTION_NAME[] = "ADC Input Portion Value";

// Declare global variables and class objects here
// No objects are created in this file

// Declare functions and procedures here
//const char* convIntToCharArray(int inputInt);

/**
	@brief ...
	@param radius ...
	@return ...
*/
// gradient finders
// float scaleAdcFloatInput(unsigned rawAdcInput, int AdcInputBitRes, float outputValueMin, float outputValueMax);
float scaleAdcFloatInput(int rawAdcInput, int AdcInputBitResMax, int AdcInputBitResMin, float outputValueMax, float outputValueMin);

//inputValueType printDebugValue(bool debugMode, String valueName, inputValueType inputValue, String valueUnit);

/**
	@brief ...
	@param radius ...
	@return ...
*/
//void printFloatDebugValue(bool debugMode, String valueName, float inputFloatValue, String valueUnit);
//
//void printIntDebugValue(bool debugMode, String valueName, int inputIntValue, String valueUnit);
//
//void printBoolDebugValue(bool debugMode, String valueName, bool inputBoolValue);

#endif
