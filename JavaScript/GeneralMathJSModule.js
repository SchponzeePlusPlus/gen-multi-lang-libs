// https://medium.com/it-developers/4-ways-to-create-a-custom-object-in-javascript-6f5e67d57500
// define a constructor
function defineTwoDimCartesianCoordinatesV000(x, y)
{
	// attribute x
	this.x = x;
	// attribute y
	this.y = y;
//	this.show=function()
//	{ // method show
//	  console.log(this.x, this.y);
//	};
}

// call constructor to create custom object
let point=new defineTwoDimCartesianCoordinatesV000(3, 4);

function defineThreeDimCartesianCoordinatesV000(x, y, z)
{
	// attribute x
	this.x = x;
	// attribute y
	this.y = y;
	this.z = z;
}

function defineTwoDimComplexCoordinatesV000(i, j)
{
	// attribute x
	this.x = x;
	// attribute y
	this.y = y;
}

function defineThreeDimComplexCoordinatesV000(i, j, k)
{
	// attribute x
	this.x = x;
	// attribute y
	this.y = y;
	this.z = z;
}

function defineQuadraticEqXIntsV000(xOne, xTwo)
{
	this.xOne = xOne;
	this.xTwo = xTwo;
}

function calcYViaLinearFuncV000(grad, xVal, yInt)
{
	return ((grad * xVal) + yInt);
}

function calcXViaQuadraticFormulaV000(a, b, c)
{
	// var result;
	let result = new defineQuadraticEqXIntsV000(
		(((-(b)) + (((b ^ (2)) - 4 * a * c) ^ (1 / 2))) / (2 * a)),
		(((-(b)) - (((b ^ (2)) - 4 * a * c) ^ (1 / 2))) / (2 * a))
	);
	return result;
}

function calcSignalRmsViaSigPeakV000(sigPeak)
{
	return (sigPeak * (1 / ((2) ^ (1 / 2))));
}