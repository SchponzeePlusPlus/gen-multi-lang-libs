function calcForceViaMAV000(mass, accel)
{
	return (mass * accel);
}

//  Coursera: University of Colorado Boulder: Power Electronics Specialization: Introduction to Power Electronics Course
//  Week 2: Steady State Conversion Analysis
//Sect. 2.2 Volt-Sec Balance and the Small Ripple Approximation

//	@brief
//	i (output period) = i / m
//	@param rate_interest_nominal: The Nominal Interest Rate (NIR) is the interest rate expressed in terms of the interest payment made each period.;
//	Interest rate of same time period to Number of compounding periods in the specified output period interval measurement parameter; Unit of Measurement is a Factor to 1
//	(0.01 interest rate factor = 1 % interest rate) => i
//	@param compounding_freq_per_output_period: Number of compounding periods per the specified output period interval measurement => m
//	@return Interest rate of same time period to specified output period interval measurement parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i

// Automotive Eng-Finance Calcs

//	@brief Calculation for Fuel Consumed per Period specified
//	Period specific parameters should all refer to the same period intervals (e.g. months or years)
//	All parameter units of measurements should be same (e.g. - km per Month & L / 100 km)
//	FCPP [fuel capacity unit / 1 period unit] = 100 * DPP [distance unit / 1 period unit] * FC [fuel capacity unit / 100 period units]
//	@param distance_per_period =>
//	@param fuel_consumed_per_100_periods => 
//	@param  => 
//	@return  => 

function calcFuelConsumedCapacityPerPeriodViaDppFCp100pV000(distance_per_period, fuel_consumed_per_100_periods)
{
	return (100.0 * distance_per_period * fuel_consumed_per_100_periods);
}