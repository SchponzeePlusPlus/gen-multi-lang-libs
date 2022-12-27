//  GeneralMathProbJSModule.js

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