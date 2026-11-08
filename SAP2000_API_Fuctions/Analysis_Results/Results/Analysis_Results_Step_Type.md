# Analysis Results Step Type

The step type varies depending on the type of output case considered. For linear static and response spectrum load cases, the step type is not used.

For nonlinear static load cases, the step type is Step, Max, or Min.

For modal cases, the step type is Mode.

For buckling cases, the step type is Mode, which refers to the buckling mode number.

For linear modal history, nonlinear modal history, linear direct integration history, and nonlinear direct integration history cases, the step type is Time, Max or Min.

For steady state and power spectral density load cases, the step type is Real at Freq, Imag at Freq, Real Min, Real Max, ImagMin, Imag Max, Mag at Freq, Mag Min, Mag Max, or RMS. Freq is short for frequency, Imag is short for imaginary and Mag is short for magnitude.

For double-valued combinations, this is either Max or Min.

For moving load patterns, this is the force correspondence if it has been requested. It may be M3 Min, M3 Max, M2 Min, M2 Max, T Min, T Max, V3 Min, V3 Max, V2 Min, V2 Max, P Min, or P Max.

Release Notes

Changed nomenclature from Load Cases, Analysis Cases and Response Combinations to Load Patterns, Load Cases and Load Combinations, respectively, in version 12.00.
