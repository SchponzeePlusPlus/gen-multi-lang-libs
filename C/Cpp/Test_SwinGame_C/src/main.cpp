/*
	Program: Question 3 for COS10008 Assignment 1 (C++)
	Module: Main
	Version: 005
	Author: Tim Flamuri

	Description:
	Draws a pattern of star shapes via SwinGame with random colours
*/

// Call necessary modules/libraries via header files
#include <stdlib.h>
#include <stdio.h>
#include <iostream>
#include <vector>
#include <array>

// for integer to character array operations
#include <sstream>
#include <string.h>
#include <time.h>
#include "SwinGame.h"

// "Magic numbers" are assigned as global constants
int
	SG_WINDOW_WIDTH = 800,
	SG_WINDOW_HEIGHT = 600,
	X_INITIAL = 100,
	Y_INITIAL = 100,
	TRIANGLE_WIDTH = 10,
	TRIANGLE_HEIGHT = 10,
	STAR_GAP_LENGTH = 15,
	STAR_ROWS_MAX = 10
;

// Call a function to draw each star based on updating coordinates
void draw_star(int clr_value, double x_centre, double y_centre, int width, int height)
{
	// Define SwinGame Color variable
	color clr;
	
	// Depending on the randomly assigned value of the parameter, a colour is assigned
	switch(clr_value)
	{
		case 0:
			clr = COLOR_RED;
			break;
		case 1:
			clr = COLOR_BLUE;
			break;
		case 2:
			clr = COLOR_YELLOW;
			break;
		case 3:
			clr = COLOR_GREEN;
			break;
	}
	
	/*
		The star shape is formed by 4 different triangles with the same width
		and height. They are paired into two diamond formations, one is vertically
		aligned whilst the other is horizontal, and overlap in the centre area.
		The x and y coordinates passed in as parameters are the centre coordinates
		of each star, and the width and height parameters are the width and height
		of each triangle. Each fill_triangle() procedure is parametrised with
		the necessary calculations based of the parameters to state the positions
		of the 12 coordinates required to draw the star. All 4 triangle procedures
		pass the same colour variable assigned according to the random number
		passed into this procedure.
	*/

	// Draw the top triangle
	fill_triangle(clr, (x_centre - (width / 2)), y_centre,
		(x_centre + (width / 2)), y_centre, x_centre, (y_centre - height));
	// Draw the left triangle
	fill_triangle(clr, (x_centre - height), y_centre, x_centre,
		(y_centre - (width / 2)), x_centre, (y_centre + (width / 2)));
	// Draw the bottom triangle
	fill_triangle(clr, (x_centre - (width / 2)), y_centre,
		(x_centre + (width / 2)), y_centre, x_centre, (y_centre + height));
	// Draw the right triangle
	fill_triangle(clr, (x_centre + height), y_centre, x_centre,
		(y_centre - (width / 2)), x_centre, (y_centre + (width / 2)));
}

// Call a procedure to initially set up the SwinGame window
void setup_gui()
{
	// Basic setup procedures for a SwinGame window
	open_graphics_window("Shape", SG_WINDOW_WIDTH, SG_WINDOW_HEIGHT);
	clear_screen(COLOR_WHITE);
}

// A procedure designed to prototype the basic logic of the program
void command_line_test(void)
{
	int
		p,
		n = 1
	;

	while (n <= 10)
	{
		int p = 1;
		while (p <= n)
		{
			printf("*");
			p++;
		}
		printf("\n");
		n++;
	}

}

// The main procedure
int main()
{
	
	int
		// Define and assign the variable that counts the rows
		n = 1,
		// Define and assign the variable that counts the stars in each row
		p = 1,
		/*
			Define and assign the variables that update the position of the
			current star to be drawn. Initially the first star is drawn
			at a specified coordinate assigned by the global constants.
		*/
		x = X_INITIAL,
		y = Y_INITIAL
	;

	// Call a procedure to initialise SwinGame window
	setup_gui();
	
	// Initialises random number generator
	srand(time(NULL));
	do
	{
		/*
			Use a conditional loop to ensure current row is within desired
			range.
		*/
		while (n <= STAR_ROWS_MAX)
		{
			/*
				The pattern or arrangement of the stars is logical. The number
				of stars in each row is equal to the row number, for example,
				row 1 has 1 star, row 10 has 10 stars.

				Use a conditional loop to ensure current star in the row is within desired
				range, which depends on the row number.
			*/
			while(p <= n)
			{
				/*
					Use a procedure to draw each star based on a random
					number that will represent the colour, the current
					coordinates of the centre x and y position, and
					the global constant values representing the width
					and height of each triangle.
				*/
				draw_star(rand() % 4, x, y,
					TRIANGLE_WIDTH, TRIANGLE_HEIGHT);
				// Count to next star in row
				p++;
				// Update x position to next star in row
				x += TRIANGLE_WIDTH + STAR_GAP_LENGTH;
			}
			// Count to next row
			n++;
			// Clear current star in row back to 1
			p = 1;
			// Clear x position back to default
			x = X_INITIAL;
			// Update y position to next row
			y += TRIANGLE_HEIGHT + STAR_GAP_LENGTH;
		}
		// Refresh screen with new objects
		refresh_screen(60);
		delay(100);
	}
	// Program is active until the SwinGame window is closed
	while (window_close_requested() == false);
	return 0;
}