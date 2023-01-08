// GeneralCalcJSModule.js

function calcMinValBetweenTwoValsV000(value_one, value_two)
{
	let result = 0;

	if (value_one <= value_two)
	{
		//  block of code to be executed if condition1 is true
		result = value_one;
	}
	else if (value_one > value_two)
	{
		//  block of code to be executed if the condition1 is false and condition2 is true
		result = value_two;
	}
	else
	{
		//  block of code to be executed if the condition1 is false and condition2 is false
		result = 0;
	}

	return result;
}