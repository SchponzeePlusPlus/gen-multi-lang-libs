#include <stdio.h>
#include "SwinGame.h"

void draw_calc () {
	clear_screen(COLOR_WHITE);

	draw_line(COLOR_BLACK, 50, 550, 350, 550);
	draw_line(COLOR_BLACK, 50, 425, 350, 425);
	draw_line(COLOR_BLACK, 50, 300, 350, 300);
	draw_line(COLOR_BLACK, 50, 175, 350, 175);
	draw_line(COLOR_BLACK, 50, 50, 350, 50);

	draw_line(COLOR_BLACK, 50, 550, 50, 50);
	draw_line(COLOR_BLACK, 125, 550, 125, 50);
	draw_line(COLOR_BLACK, 200, 550, 200, 50);
	draw_line(COLOR_BLACK, 275, 550, 275, 50);
	draw_line(COLOR_BLACK, 350, 550, 350, 50);

	draw_simple_text("+", COLOR_BLACK, 60, 540);
	draw_simple_text("-", COLOR_BLACK, 135, 540);
	draw_simple_text("X", COLOR_BLACK, 210, 540);
	draw_simple_text("/", COLOR_BLACK, 285, 540);
	draw_simple_text("9", COLOR_BLACK, 60, 415);
	draw_simple_text("0", COLOR_BLACK, 135, 415);
	draw_simple_text("=", COLOR_BLACK, 210, 415);
	draw_simple_text("Clear", COLOR_BLACK, 285, 415);
	draw_simple_text("5", COLOR_BLACK, 60, 290);
	draw_simple_text("6", COLOR_BLACK, 135, 290);
	draw_simple_text("7", COLOR_BLACK, 210, 290);
	draw_simple_text("8", COLOR_BLACK, 285, 290);
	draw_simple_text("1", COLOR_BLACK, 60, 165);
	draw_simple_text("2", COLOR_BLACK, 135, 165);
	draw_simple_text("3", COLOR_BLACK, 210, 165);
	draw_simple_text("4", COLOR_BLACK, 285, 165);

	draw_simple_text("Operand 1", COLOR_BLACK, 530, 165);
	draw_simple_text("Operand 2", COLOR_BLACK, 530, 290);
	draw_simple_text("Operator", COLOR_BLACK, 530, 415);
	draw_simple_text("Result", COLOR_BLACK, 530, 540);

	draw_line(COLOR_BLACK, 600, 550, 750, 550);
	draw_line(COLOR_BLACK, 600, 425, 750, 425);
	draw_line(COLOR_BLACK, 600, 300, 750, 300);
	draw_line(COLOR_BLACK, 600, 175, 750, 175);
	draw_line(COLOR_BLACK, 600, 50, 750, 50);

	draw_line(COLOR_BLACK, 600, 550, 600, 50);
	draw_line(COLOR_BLACK, 750, 550, 750, 50);

	refresh_screen(60);
}

void run_calc() {
	do
	{
		process_events();
		draw_calc();
	}
	while (window_close_requested() == false);
}

void setup_gui() {
	open_graphics_window("Simple Calculator", 800, 600);
	clear_screen(ColorWhite);
	refresh_screen(60);
	delay(2000);
}

/*
int old_main()
{
	int count1=0, value=0, number1=0, operation=0;
	
	printf("Enter the first number, one digit at a time ... or enter -1 to stop \n");
	scanf("%d", &value);
	while(value != -1)
	{
		if(count1==0) //this condition verifies whether the input is the first digit
		{
			if(value!=0) //this condition verifies that the first digit is not a zero
			{
				number1=value; //assign the first digit to number1
				count1+=1; // increments the number of digit count by 1
			}
		}
		else
		{
			number1=number1*10+value; // new digit is added to the existing number
			count1+=1;
		}	
		
		printf("Enter the next digit of first number or 0-1 to stop ...\n");
		scanf("%d", &value);
}
printf("the first number = %d \n",number1);
//code for the selection of operation

	printf("Enter the number corresponding to the operation type ...\n");
	printf("1 Addition \n"); printf("2 Subtraction \n"); printf("3
	Multiplication \n"); printf("4 Division \n");
	scanf("%d",&operation);
	
	
return 0;
}
*/

int main()
{
	setup_gui();
	run_calc();
	
	return 0;
}