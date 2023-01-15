//  GeneralMathProbJSModule.js

var EventProbabilityStateV000;
(function (EventProbabilityStateV000) {
	EventProbabilityStateV000[EventProbabilityStateV000["MUTUALLY_EXCLUSIVE_EPS"] = 0] = "MUTUALLY_EXCLUSIVE_EPS";
	EventProbabilityStateV000[EventProbabilityStateV000["INDEPENDENT_EPS"] = 1] = "INDEPENDENT_EPS";
	EventProbabilityStateV000[EventProbabilityStateV000["NULL_EPS"] = 2] = "NULL_EPS";
})(EventProbabilityStateV000 || (EventProbabilityStateV000 = {}));

const EventProbabilityStateV001 = Object.freeze
({
	NULL_EPS: Symbol(0),
	MUTUALLY_EXCLUSIVE_EPS: Symbol(1),
	INDEPENDENT_EPS: Symbol(2)
})

// mutually exclusive, aka disjoint events
const MUTUALLY_EXCLUSIVE_EPS_STRING_VAL = "MUTUALLY_EXCLUSIVE";
const INDEPENDENT_EPS_STRING_VAL = "INDEPENDENT";

function assignEventProbabilityStateV000EnumFromStringV000(input_str)
{
	let result = EventProbabilityStateV001.NULL_EPS;
	switch(input_str)
	{
		case MUTUALLY_EXCLUSIVE_EPS_STRING_VAL:
			result = EventProbabilityStateV001.MUTUALLY_EXCLUSIVE_EPS;
		case INDEPENDENT_EPS_STRING_VAL:
			result = EventProbabilityStateV001.INDEPENDENT_EPS;
		default:
			result = EventProbabilityStateV001.NULL_EPS;
	}

	return result;
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 03, Video x - Addition Rules for Probability
// P(A u B) = P(A) + P(B)
function addProbEventsMutExclusViaPaPbV000(prob_a, prob_b)
{
	return (prob_a + prob_b);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 03, Video x - Addition Rules for Probability
// P(A u B) = P(A) + P(B) - P(A&B)
function addProbEventsIndependentViaPaPbPanbV000(prob_a, prob_b, prob_a_and_b)
{
	return (prob_a + prob_b - prob_a_and_b);
}

function addProbEventsIndependentViaPaPbPanbEpsV000(prob_a, prob_b, prob_a_and_b, eps)
{
	let result = 0;

	switch(eps)
	{
		case EventProbabilityStateV001.MUTUALLY_EXCLUSIVE_EPS:
			result = addProbEventsMutExclusViaPaPbV000(prob_a, prob_b);
		case EventProbabilityStateV001.INDEPENDENT_EPS:
			result = addProbEventsIndependentViaPaPbPanbV000(prob_a, prob_b, prob_a_and_b);
		default:
			result = 0;
	}

	return result;
}

//  PLEASE NOTE THAT THIS FUNCTION CURRENTLY DOES NOT WORK IN A GAScript for a Google Sheet (https://docs.google.com/spreadsheets/d/1jWfozhzJdyaWn_0ZAY6LvFL4bO2ue-QeE8RuyQ09Zq0/edit#gid=0) - PLEASE TROUBLESHOOT FIRST
function addProbEventsIndependentViaPaPbPanbEpsstrV000(prob_a, prob_b, prob_a_and_b, eps_str)
{
	return addProbEventsIndependentViaPaPbPanbEpsV000(prob_a, prob_b, prob_a_and_b, assignEventProbabilityStateV000EnumFromStringV000(eps_str));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 03, Video x - Conditional Probability
//  Formula for probabiity of event B when event A has occured:
//  P(B | A) = P(A&B) + P(A)
function calcConditionalProbEventViaPanbPbV000(prob_a_and_b, prob_a)
{
	return (prob_a_and_b / prob_a);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 03, Video x - Multiplicative Rules
//  Special Multiplication Rule (Independent Events)
//  P(A n B) = P(A) * P(B)
function multiplyProbEventsIndependentViaPaPbV000(prob_a, prob_b)
{
	return (prob_a * prob_b);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 03, Video x - Multiplicative Rules
//  General Multiplication Rule (Mutually Exclusive Events)
//  P(A n B) = P(A) * P(B | A)
function multiplyProbEventsMutExclusViaPaPbaV000(prob_a, cond_prob_b_after_a)
{
	return (prob_a * cond_prob_b_after_a);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 2 - Central Limit Theorem
//  sigma ^ 2 = sigma_x-bar ^ 2 = sigma_x ^ 2 / n
function calcVarianceViaSigmaxNV000(sigma_x, n)
{
	return (sigma_x ** 2 / n);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 2 - Central Limit Theorem
//  Standard Deviation: sigma = sigma_x-bar = sigma_x / sqrt(n)
function calcStdDevViaSigmaxNV000(sigma_x, n)
{
	return (sigma_x / (n ** (1 / 2)));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 1 - Combinations & Permutations
function calcPermutationsViaNRV000(n_obj_sets, r_objs)
{
	// nPr; ((n!) / (n - r)!)
	return (calcNFactorialV000(n_obj_sets) / calcNFactorialV000(n_obj_sets - r_objs));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 1 - Combinations & Permutations
// nCr; ((n!) / [r! * (n - r)!])
function calcCombinationsViaNRV000(n_obj_sets, r_objs)
{
	return (calcNFactorialV000(n_obj_sets) / (calcNFactorialV000(r_objs) * calcNFactorialV000(n_obj_sets - r_objs)));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 3 - Central Limit Theorem Applications
//  Standard Error of the mean (aka Std Deviation of the mean) = sigma / sqrt(n)
//	@param ppltn_std_dev Population parameter: standard deviation => sigma
//	@param smpl_n Sample parameter: Sample Size n => n
function calcStdErrOTMeanViaSigmaNV000(ppltn_std_dev, smpl_n)
{
	return (ppltn_std_dev / (smpl_n ** (1 / 2)));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 3 - Central Limit Theorem Applications
//  Confidence Intervals
//	...

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 05 (Statistical Distributions), Video x - Binomial Distribution Calculations
//  Binomial Formula
//	@brief
//	P(x) = ((n!) / (x! * ((n - x)!))) * p ^ (x) * (1 - p) ^ (n - x)
//	@param
//	@return
function calcBinomialFormulaViaNumsuccNumtrialProbSuccV000(num_success, num_trials, prob_success)
{
	return (((calcNFactorialV000(num_trials)) / (calcNFactorialV000(num_success) * (calcNFactorialV000(num_trials - num_success)))) * prob_success ** (num_success) * (1 - prob_success) ** (num_trials - num_success));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 05 (Statistical Distributions), Video x - 
//  Rate for the problem
//	@brief
//	Lambda = n * p
//	@param
//	@return
function calcProblemRateViaNumsuccNumtrialV000(num_success, num_trials)
{
	return (num_success * num_trials);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 05 (Statistical Distributions), Video x - 
//  Poisson Formula
//	@brief
//	P(x) = ((e ^ (-lambda) * lambda ^ (x)) / (x!))
//	@param
//	@return
function calcPoissonFormulaViaNumsuccProblemrateV000(num_success, problem_rate)
{
	return ((Math.exp(-1.0 * problem_rate) * problem_rate ** (num_success)) / (calcNFactorialV000(num_success)));
}

//	MTH20004: Maths 3
//	Student Notes Pg. 198
//  Probability Density Function
//	Normal Distribution ??
//	f(x) = (1 / (sigma * sqrt(2 * Pi))) * e ^ (-(1 / 2) * ((x - mu) / sigma) ^ 2)
//	-ve inf < x < +ve inf
// doesnt work properly?!!
// might have excessive brackets - could remove some in future
function calcProbDensityFuncContRvViaSigmaMuXV000(mean, std_dev, x)
{
	return ((1.0 / (std_dev * ((2.0 * Math.PI) ** (1.0 / 2.0)))) * Math.exp(-1.0 * (1.0 / 2.0) * (((x - mean) / std_dev) ** (2.0))));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 05 (Statistical Distributions), Video x - 
//  Z-Score
//	@brief
//	Z = (x - mu) / sigma
//	@param
//	@return
function calcZScoreViaXMeanStddevV000(x, mean, std_dev)
{
	return ((x - mean) / std_dev);
}

// Probability between range of values of normal distribution
// function to work out area under graph; ugly method ?
function calcProbNormDistrRangeViaMuSigmaMinxMaxxV000(mean, std_dev, min_x, max_x)
{
	// let integration_smpling_intrvl = Number.MIN_VALUE;
	//	let integration_smpling_intrvl = 1.0;
	let integration_smpling_intrvl = 0.001;
	let i = 0;
	let x = 0;
	let range_x = 0;
	let integration_cnt_lim = 0;
	let result = 0;
	let prob_density_x = 0;
	let prob_density_x_prev = 0;

	range_x = max_x - min_x;
	integration_cnt_lim = Math.round(range_x / integration_smpling_intrvl);

	/* 
	x = min_x;
	
	for (i = 0; i < integration_cnt_lim; i++)
	{
		result += calcProbDensityFuncContRvViaSigmaMuXV000(mean, std_dev, x);
		x += integration_smpling_intrvl;
	}
	*/

	x = min_x;
	
	prob_density_x_prev = calcProbDensityFuncContRvViaSigmaMuXV000(mean, std_dev, x);

	x += integration_smpling_intrvl;

	for (i = 1; i < integration_cnt_lim; i++)
	{
		prob_density_x = calcProbDensityFuncContRvViaSigmaMuXV000(mean, std_dev, x);
		
		result += (calcMinValBetweenTwoValsV000(prob_density_x, prob_density_x_prev) * integration_smpling_intrvl) + ((1 / 2) * integration_smpling_intrvl * Math.abs(prob_density_x - prob_density_x_prev));

		x += integration_smpling_intrvl;
		prob_density_x_prev = prob_density_x;
	}

	return result;
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 05 (Statistical Distributions), Video x - 
//  t-distribution
//	@brief
// t = (x) / sqrt(y / k)
//	@param k Degrees of Freedom
//	@return
function calcTDistributionViaRvndRvcsDofV000(random_var_norm_dist, random_var_chi_sqr_dist, dof)
{
	return ((random_var_norm_dist) / ((random_var_chi_sqr_dist / dof) ** (1 / 2)));
}