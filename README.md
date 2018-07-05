# excel_xcov by Nury Kim, IBS Korea, tesdar@ibs.re.kr
This macro provides MATLAB xcov calculation + LOWESS smoothing for two dataset

The xcov is different from cross covariance, which is a cross correlation of 'inverted' one signal and the other. 
The macro follows as described in MATLAB xcov function; average or mean adjusted cross correlation.

To use the macro, you need two dataset in two columns, starting from A1 and B1
The macro will use from column C to J to show the results, so empty or make a copy of the area before use

Run the macro 'xCov' and will give you results:
 C and D are average adjusted data of A and B respectively
 E is shifted version of C, by the length N - 1 where N is the number of data
 F is the shifted timeline, and the distance is defined as 1
 G is the xcov result
 I and J are for LOWESS smoothing, I for timeline at 0.01 precision, J is smoothed result
 
