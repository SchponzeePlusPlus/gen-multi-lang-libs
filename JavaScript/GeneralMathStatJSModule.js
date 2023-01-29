//	GeneralMathStatJSModule.js

function calcAverageValViaValsarrV000(...vals_arr)
{
	let vals_arr_sum = 0;
	let i = 0;

	for (i = 0; i < vals_arr.length; i++)
	{
		vals_arr_sum += vals_arr[i];
	}

	return (vals_arr_sum / vals_arr.length);
}