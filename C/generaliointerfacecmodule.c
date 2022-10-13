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

float scaleAdcFloatInputV001(int )