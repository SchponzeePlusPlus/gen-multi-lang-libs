#ifndef INCLUDES_GENERALDATATYPECMODULE_H_
#define INCLUDES_GENERALDATATYPECMODULE_H_


// for future use in turning off the ldrs
const float LDR_REF_R_AVE_MIN = 0.05;

// Declare enumerators
// Declare an enumerator that specifies the evaluation after comparing two opposing LDR sets

enum LDRSet1to2CompareState
{
	SET_GREATER,
	SET_LOWER,
	SET_EQUAL,
	ERROR_STATE
};

// Define records or structs
// Declare a struct that contains all the ADC values from the LDR reference resistor inputs
struct LDRRefRRawRecord
{
	int NEBits;
	int SEBits;
	int SWBits;
	int NWBits;
};

#endif