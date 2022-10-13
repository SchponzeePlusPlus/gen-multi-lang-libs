// Call necessary modules/libraries via header files

// ...

int calcRangeIntV000(int max, int min)
{
	return (max - min);
}

float calcYViaLinearFuncFloatV000(float grad, float x_val, float y_int)
{
	float result = ((grad * x_val) + y_int);
	return result;
}