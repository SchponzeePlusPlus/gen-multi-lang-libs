//	GeneralMathStatJSModule.js

function calcAverageValViaValsarrV000(vals_arr)
{
	let vals_arr_sum = 0;
	let i = 0;

	for (i = 0; i < vals_arr.length; i++)
	{
		vals_arr_sum += vals_arr[i];
	}

	return (vals_arr_sum / vals_arr.length);
}

function calcAverageValViaValsarrV001(vals_arr)
{
	return (sumValsViaValsarrV000(vals_arr) / vals_arr.length);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//  Grand Average; Mean of Means
//	@brief
//	X_double-bar = Average(X_bar values)
//	@param
//	@param
//	@param
//	@return Grand Average; Mean of Means => X_double-bar
function calcGrandAvgValViaAvgarrV000(avgs_arr)
{
	return (calcAverageValViaValsarrV001(avgs_arr));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//	@brief
//	R_bar = Average(R values)
//	@param
//	@param
//	@param
//	@return Average Range => R_bar
function calcRngeAvgValViaRngearrV000(rnges_arr)
{
	return (calcAverageValViaValsarrV001(rnges_arr));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//  X-bar Chart Upper Control Limit
//	@brief
//	UCL_X-bar = X_double-bar + A_2 * R_bar
//	@param
//	@param ctrl_lim_const_A_two => A_2
//	@param
//	@return
function calcUCLXBarViaGrandavgAtwoRngeavgV000(grand_avg, ctrl_lim_const_A_two, rnge_avg)
{
	return (grand_avg + ctrl_lim_const_A_two * rnge_avg);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//	@brief
//	LCL_X-bar = X_double-bar - A_2 * R_bar
//	@param
//	@param ctrl_lim_const_A_two => A_2
//	@param
//	@return X-bar Chart Lower Control Limit
function calcLCLXBarViaGrandavgAtwoRngeavgV000(grand_avg, ctrl_lim_const_A_two, rnge_avg)
{
	return (grand_avg - ctrl_lim_const_A_two * rnge_avg);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//	@brief
//	UCL_R = D_4 * R_bar
//	@param
//	@param ctrl_lim_const_D_four => D_4
//	@param
//	@return R Chart Upper Control Limit
function calcUCLRBarViaDfourRngeavgV000(ctrl_lim_const_D_four, rnge_avg)
{
	return (ctrl_lim_const_D_four * rnge_avg);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//	@brief
//	LCL_R = D_3 * R_bar
//	@param
//	@param ctrl_lim_const_D_three => D_3
//	@param
//	@return R Chart Lower Control Limit
function calcUCLRBarViaDfourRngeavgV000(ctrl_lim_const_D_three, rnge_avg)
{
	return (ctrl_lim_const_D_three * rnge_avg);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 2 - Designing
//		Process Performance
//	@brief
//	<math equation>
//	@param
//	@param
//	@param
//	@return <function title or result>