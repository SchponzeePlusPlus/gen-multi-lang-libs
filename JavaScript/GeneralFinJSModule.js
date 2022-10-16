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