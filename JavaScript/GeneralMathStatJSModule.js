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
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 3 - Analyzing
//		Process Capability
//	@brief
//	σ_circumflex = R_bar / d_2
//	@param =>
//	@param =>
//	@param =>
//	@return estimated standard dev => σ_circumflex
function calcEstStdDevViaRngeavgDtwoV000(rnge_avg, ctrl_lim_const_D_two)
{
	return (rnge_avg / ctrl_lim_const_D_two);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 3 - Analyzing
//		Process Capability
//	@brief
//	C_pk = Min(Z_U, Z_L) / 3
//	@param Z_U aka std_devs_between_proc_avg_and_usl : number of standard deviations between the process average and the upper specification
//		limit => Z_U
//	@param => Z_L
//	@param
//	@return Process Performance "Capability Index" => C_pk
function calcProcPerfCapbIndexViaZuZlV000(Z_U, Z_L)
{
	return (calcMinValBetweenTwoValsV000(Z_U, Z_L) / 3);
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 3 - Analyzing
//		Process Capability
//	@brief
//	C_p = Tolerance Range / (6 * σ_circumflex)
//	@param tolerance range =>
//	@param estimated standard dev =>
//	@param =>
//	@return Capability Index => C_p
function calcCapbIndexViaZuZlV000(tol_rnge, est_std_dev)
{
	return (tol_rnge / (6 * est_std_dev));
}

//  Coursera Kennesaw State University Six Sigma Green Belt Specialisation (SSGBSpec) Course 02: Adavanced Define and Measure Phases (ADMP)
//  Week 07 (Process and Performance Capability & Exploratory Data Analysis), Part 1 - Process and Performance Capability, Video 3 - Analyzing
//		Process Capability
//	@brief
//	<math equation>
//	@param =>
//	@param =>
//	@param =>
//	@return <function title or result> =>