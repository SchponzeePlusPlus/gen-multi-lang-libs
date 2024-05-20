/*
test
*/

# test in-line

#	Electrical Power
def calcPwrWViaVoltVCurrA(voltV, currA):
	return (voltV * currA)

#	Rotational Power
def calcPwrWViaTrqNmAngVelradps(trqNm, angVelRadPerSec):
	return (trqNm * angVelRadPerSec)

#	Mecahnical / Hydraulic Power
#	not sure what the units should be used to calculate Hydraulic Power in Watts
def calcPwrWViaPressXFloRteY(pressX, floRteY):
	return (pressX * floRteY)
