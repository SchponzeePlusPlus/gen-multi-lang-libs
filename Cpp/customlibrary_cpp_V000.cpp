/*
 ============================================================================
 * @file    customlibrary.cpp (180.ARM_Peripherals/Sources/level.cpp)
 * @brief   Basic C++ dataplotting module
 *
 *  Created on: 15/09/2018
 *      Author: Sponza
 *		File Version: V000
 *		Purpose of File:
 *
 *		The level.cpp file is the module responsible for communicating with
 *		the accelerometer and LCD, the procedures and functions report/scale
 *		values from the accelerometer and then create drawings onto the LCD.
 *		They are independently called from main() in the main.cpp file.
 ============================================================================
 */

/*
	Header files allow modules to access C++ libraries and other modules
	apart of the project.
*/
#include <stdlib.h>
#include <iostream>
#include <vector>
#include <array>

// for integer to character array operations
#include <sstream>
#include <string.h>

const char* convIntToCharArray(int inputInt) {
	std::stringstream tempCharArray;
		tempCharArray << (inputInt);

			std::string charArrayString = tempCharArray.str();
			const char* outputCharArray = charArrayString.c_str();
	return outputCharArray;
}

// gradient finders