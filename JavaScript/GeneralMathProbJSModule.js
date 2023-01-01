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

function addProbEventsIndependentViaPaPbPanbEpsstrV000(prob_a, prob_b, prob_a_and_b, eps_str)
{
    return addProbEventsIndependentViaPaPbPanbEpsV000(prob_a, prob_b, prob_a_and_b, assignEventProbabilityStateV000EnumFromStringV000(eps_str));
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