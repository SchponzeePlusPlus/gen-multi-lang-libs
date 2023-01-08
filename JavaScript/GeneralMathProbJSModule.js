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
	return (sigma_x ^ 2 / n);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 2 - Central Limit Theorem
//  Standard Deviation: sigma = sigma_x-bar = sigma_x / sqrt(n)
function calcStdDevViaSigmaxNV000(sigma_x, n)
{
	return (sigma_x / (n ^ (1 / 2)));
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
function calcStdErrOTMeanViaSigmaNV000(sigma, n)
{
	return (sigma / (n ^ (1 / 2)));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 04, Video 3 - Central Limit Theorem Applications
//  Confidence Intervals
//	...

//	MTH20004: Maths 3
//	Student Notes Pg. 198
//  Probability Density Function
//	f(x) = (1 / (sigma * sqrt(2 * Pi))) * e ^ (-(1 / 2) * ((x - mu) / sigma) ^ 2)
//	-ve inf < x < +ve inf
// doesnt work properly?!!
function calcProbDensityFuncContRvViaSigmaMuXV000(mean, std_dev, x)
{
	return ((1 / (std_dev * (2 * Math.PI) ^ (1 / 2))) * Math.exp(-(1 / 2) * ((x - mean) / std_dev) ^ 2));
}