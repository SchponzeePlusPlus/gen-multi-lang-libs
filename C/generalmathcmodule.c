// Call necessary modules/libraries via header files

// ...

float calcYViaLinearFuncFloatV000(float grad, float x_val, float offset)
{
	float result = ((grad * x_val) + offset);
	return result;
}