//  GeneralMathProbJSModule.js

//  Coursera SSGBSpec_02_ADMP
function calcPermutationsViaNRV000(n_obj_sets, r_objs)
{
    // nPr; ((n!) / (n - r)!)
    return (calcNFactorialV000(n_obj_sets) / calcNFactorialV000(n_obj_sets - r_objs));
}