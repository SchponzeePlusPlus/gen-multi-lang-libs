function calcForceViaMAV000(mass, accel)
{
	return (mass * accel);
}

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