// Call necessary modules/libraries via header files

// ...

float calcYViaLinearFuncFloatV000(float grad, float x_val, float y_int)
{
	float result = ((grad * x_val) + y_int);
	return result;
}