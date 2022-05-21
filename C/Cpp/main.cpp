/*
	Program: Test pointers
	Module: Main
	Author: Leonard Sponza
	100588917

	Description:
	 
*/

// Call necessary modules/libraries via header files
#include <algorithm>
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
#include <time.h>

// Use the C++ standard namespace which includes cin and cout
using namespace std;

// Declare enumerators; none declared


// Declare records or structs

// declare defines

// Declare global constants, "magic numbers" are assigned as global constants

// Declare the main function
int main ()
{
	// Declare variables to store numbers
	// Declare a local variable to store result from function
	
	cout << "Welcome to the program:" << endl;
	int var_one = 10;
	int var_two = 20;
	// int y = &var;
	// Initialise a pointer with a pointer declaration
	int *yPtr = &var_one;
	// can also be declared without assignment
	int *xPtr;
	// pointers are assigned address
	xPtr = &var_two;
	// a dereference operator is used to identify the variable stored
	// in the address
	int value_var_one = *yPtr;

	cout << "contents of var_one are " << var_one << endl;
	cout << "address of var_one is " << &var_one << endl;
	cout << "contents of var_two are " << var_two << endl;
	cout << "address of var_two is " << &var_two << endl;
	//cout << "contents of y are " << y << endl;
	//cout << "variable whose address is stored at y is " << *y << endl;
	cout << "contents of yPtr are " << yPtr << endl;
	cout << "variable whose address is stored at yPtr is " << *yPtr << endl;
	cout << "contents of xPtr are " << xPtr << endl;
	cout << "variable whose address is stored at xPtr is " << *xPtr << endl;

	cout << "contents of value_var_one are " << value_var_one << endl;

	// assign a new value into a dereferenced pointer to store
	// the value in the pointer's address 
	*yPtr = 15;
	xPtr = yPtr;

	cout << "contents of var_one are " << var_one << endl;
	cout << "address of var_one is " << &var_one << endl;
	cout << "contents of var_two are " << var_two << endl;
	cout << "address of var_two is " << &var_two << endl;
	//cout << "contents of y are " << y << endl;
	//cout << "variable whose address is stored at y is " << *y << endl;
	cout << "contents of yPtr are " << yPtr << endl;
	cout << "variable whose address is stored at yPtr is " << *yPtr << endl;
	cout << "contents of xPtr are " << xPtr << endl;
	cout << "variable whose address is stored at xPtr is " << *xPtr << endl;
	
	
	// Terminate main thereby the program
	return 0;
}