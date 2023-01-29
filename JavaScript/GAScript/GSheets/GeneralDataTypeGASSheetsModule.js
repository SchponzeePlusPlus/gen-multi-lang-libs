//	GeneralDataTypeGASSheetsModule.js

//	Google Sheets Range
//	https://webapps.stackexchange.com/questions/10629/how-to-pass-a-range-into-a-custom-function-in-google-spreadsheets/58179#58179

//	@brief downwards along column first (row by row), then goes rightwards along first row to next column and then reads downwards along column again
//	PLEASE UPDATE THE NAMING CONVENTIONS IN THIS FORMULA WHEN POSSIBLE
//	@param =>
//	@param =>
//	@param =>
//	@return <function title or result> =>
function convertRngeValsToOneDimArrDownColThenRightAlongRowsViaRngeV000(range_cell_vals_two_dim_arr)
{
	let result_length = 0;
	//  rows
	let i = 0;
	//  cols
	let j = 0;
	let k = 0;
	let result = new Array();

	// "number of rows with range.length and the number of columns with range[0].length"
	result_length = range_cell_vals_two_dim_arr.length * range_cell_vals_two_dim_arr[0].length;

	result = new Array(result_length).fill(0);

	for (j = 0; j < range_cell_vals_two_dim_arr[0].length; j++)
	{
		for (i = 0; i < range_cell_vals_two_dim_arr.length; i++)
		{
			result[k] = range_cell_vals_two_dim_arr[i][j];
			k += 1;
		}
	}

	return result;
}