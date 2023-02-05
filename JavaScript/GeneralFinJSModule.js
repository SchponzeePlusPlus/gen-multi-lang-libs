//	GeneralFinJSModule.js

/// ...

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 1: Topic 1 - Introduction to Financial Management
//	@brief
//	NPV = PV of future cashflows - Outlay
//	@param =>
//	@param =>
//	@param =>
//	@return Net Present Value => NPV

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	Ross, SA, Jordan, BD &amp; Westerfield, R 2010, Fundamentals of Corporate Finance, 9th ed, McGraw-Hill Irwin, Boston.
//	ISBN 978-0-07-338239-5 (standard edition)
//	pg. 138
//	@brief
//	FVF = (1 + i) ^ t
//	@param =>
//	@param rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
//	@param periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
//	@return Future Value Factor => FVF
function calcFutureValFactorCompoundIRViaIntPerV000(rate_interest, periods_num)
{
	return ((1.0 + rate_interest) ** (periods_num));
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	Ross, SA, Jordan, BD &amp; Westerfield, R 2010, Fundamentals of Corporate Finance, 9th ed, McGraw-Hill Irwin, Boston.
//	ISBN 978-0-07-338239-5 (standard edition)
//	pg. 138
//	@brief
//	FVF = (1 + i) ^ t
//	@param =>
//	@param rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
//	@param periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
//	@return Future Value Factor => FVF
function calcPresentValFactorCompoundIRViaIntPerV000(rate_interest, periods_num)
{
	// return (1.0 / ((1.0 + rate_interest) ** (periods_num)));

	return (1.0 / calcFutureValFactorCompoundIRViaIntPerV000(rate_interest, periods_num));
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	@brief
//	FV = PV * (1 + i) ^ t
//	@param =>
//	@param rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
//	@param periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
//	@return Future Value of Lump Sum => FV
function calcFutureValLumpSumCompoundIRViaPresvallsIntPerV000(pres_val_lump_sum, rate_interest, periods_num)
{
	return (pres_val_lump_sum * ((1.0 + rate_interest) ** (periods_num)));
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	@brief
//	PV = FV / ((1 + i) ^ t)
//	@param =>
//	@param rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
//	@param periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
//	@return Present Value of Lump Sum => PV
function calcPresentValLumpSumCompoundIRViaFuturevallsIntPerV000(future_val_lump_sum, rate_interest, periods_num)
{
	return (future_val_lump_sum / ((1.0 + rate_interest) ** periods_num));
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	@brief
//	i = ((FV / PV) ^ (1 / t)) - 1
//	@param =>
//	@param Present Value of Lump Sum => PV
//	@param periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
//	@return rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
function calcIntRateCompoundedViaFuturevallsPresvallsPerV000(future_val_lump_sum, pres_val_lump_sum, periods_num)
{
	return (((future_val_lump_sum / pres_val_lump_sum) ** (1.0 / periods_num)) - 1.0);
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	@brief
//	t = (log(FV / PV)) / (log(1 + i))
//	@param =>
//	@param Present Value of Lump Sum => PV
//	@param rate_interest: Interest rate of same time period to period parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
//	@return periods_num: number of periods; unspecified time intervals that have elapsed; period intervals must match interest rate time period parameter => t
function calcPeriodsNumCompoundedViaFuturevallsPresvallsIntV000(future_val_lump_sum, pres_val_lump_sum, rate_interest)
{
	return ((calcLogViaBaseV000(10, (future_val_lump_sum / pres_val_lump_sum))) / (calcLogViaBaseV000(10, (1.0 + rate_interest))));
}

//  Swinburne University of Technology: Bachelor of Business (Finance) [BB-EngBus]: Financial Management Unit of Study (FIN20014)
//  Week 2: Topic 2A - Time Value of Money
//	@brief
//	i (output period) = i / m
//	@param rate_interest_nominal: The Nominal Interest Rate (NIR) is the interest rate expressed in terms of the interest payment made each period.;
//	Interest rate of same time period to Number of compounding periods in the specified output period interval measurement parameter; Unit of Measurement is a Factor to 1
//	(0.01 interest rate factor = 1 % interest rate) => i
//	@param compounding_freq_per_output_period: Number of compounding periods per the specified output period interval measurement => m
//	@return Interest rate of same time period to specified output period interval measurement parameter; Unit of Measurement is a Factor to 1 (0.01 interest rate factor = 1 % interest rate) => i
function calcIntRateOutFreqViaIntCompperiodsV000(rate_interest_nominal, compounding_freq_per_output_period)
{
	return (rate_interest_nominal / compounding_freq_per_output_period);
}

//	@param in_np Net Profit; financial gain
//	@param in_ti Total Investment;
//	Return on Investment (ROI) [%]
function calcROIViaNpTiV000(in_np, in_ti)
{
	return ((in_np / in_ti) * 100);
}

//	@param in_inv Investment
//	@param in_gpp Gains per Period; per month or year
//	Payback Period
function calcPaybackPeriodViaInvGppV000(in_inv, in_gpp)
{
	return (in_inv / in_gpp);
}

//	Six Sigma Yellow Belt
//	@param in_if_curr IF (Monetary Value in selected Currency); Internal Failures
//	@param in_ef_curr EF (Monetary Value in selected Currency); External Failures
//	@param in_a_curr A (Monetary Value in selected Currency); Appraisals
//	@param in_p_curr P (Monetary Value in selected Currency); Preventions
//	@param in_hidden_costs_curr Hidden Costs (Monetary Value in selected Currency)
//	@return
function calcCostOfQualityViaIfEfAPHcV000(in_if_curr, in_ef_curr, in_a_curr, in_p_curr, in_hidden_costs_curr)
{
	return (in_if_curr + in_ef_curr + in_a_curr + in_p_curr + in_hidden_costs_curr);
}

//	@param _ Input
//	@return

//	@param
//	@return

// Six Sigma Sigma Level Formula
//	@param opp_num Number of Opportunities
//	@param defect_num Number of Defects
//	@return Yield (Sigma Level)
function calcSigmaLvlViaOppDefV000(opp_num, defect_num)
{
	return (((opp_num - defect_num) / opp_num) * 100);
}