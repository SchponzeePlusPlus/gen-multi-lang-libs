// Six Sigma Sigma Level Formula
//	@param opp_num Number of Opportunities
//	@param defect_num Number of Defects
//	@return Yield (Sigma Level)
function calcSigmaLvlViaOppDefV000(opp_num, defect_num)
{
	return (((opp_num - defect_num) / opp_num) * 100);
}