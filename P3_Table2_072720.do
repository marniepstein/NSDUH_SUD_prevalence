/*************************
Project: NSDUH Tabs

Created by: Marni Epstein
Created on: March 3, 2020

Purpose: Tabs for adolescents and young adults 

Notes: Use invttail to find the CI 5th and 95% perciles of the t-model
	Use ttail to find the p-value for a 2-sided t-test
	http://www.stat.columbia.edu/~regina/W1111S09/tutorial5.pdf
	
**************************/

/***********************
SET GLOBALS - CHANGE
***********************/
*Enter user name
global user = "MEpstein"

*Enter computer drive
global drive = "D"


/****************************************************************
UPDATE DIRECTORIES - WILL UPDATE AUTOMATED BASED ON GLOBALS ABOVE
****************************************************************/
*Set directory and output folder
global boxfolder "${drive}:\Users\\${user}\Box\2020 FORE OUD\Task 1 NSDUH"
cd "${boxfolder}\Analysis\Created datasets"
gl rawdata "${boxfolder}\Data\downloaded datafiles"
gl output "${boxfolder}\Task 1 NSDUH\Analysis\Tables"

*Change if there is a more recent version
gl exceloutput "${boxfolder}\\Write ups\brief 1 prevalence\Table shells 072720.xlsx"
putexcel set "$exceloutput", modify sheet("Sheet 2")


clear
use "analytic.dta" // make sure this is the most recent created datafile


*Subset to Medicaid
keep if insurance == 2

/*******************************		
PROGRAM TO CREATE STARS 
*******************************/
capture program drop stars
program define stars
	args pval num
		
		*Add stars
		if `pval' <= 0.01 {
			gl star`num' = "***" 
		} 
		else if `pval' <= 0.05 { 
			gl star`num' = "**" 
		} 
		else if `pval' <= 0.1 { 
			gl star`num' = "*" 
		} 
		else {
			gl star`num' = "" 
		} 
end

/*******************************		
PROGRAM TO CREATE PLUSES 
*******************************/
capture program drop plus
program define plus
	args pval num
		
		*Add stars
		if `pval' <= 0.01 {
			gl plus`num' = "+++" 
		} 
		else if `pval' <= 0.05 { 
			gl plus`num' = "++" 
		} 
		else if `pval' <= 0.1 { 
			gl plus`num' = "+" 
		} 
		else {
			gl plus`num' = "" 
		} 
end


/******* Suppression Program ***************/
/**  where mean_out = mean, semean =se, nsum = sample size**/

gen rse = .
gen rselnp = .

gen suprule1a = 0
gen suprule1b = 0
gen suprule1c = 0
gen suprule2 = 0
gen suprule3 = 0

capture program drop suppressprog
program define suppressprog
	args mean_out semean nsum 
	
	/*CALCULATE THE RELATIVE STANDARD ERROR*/
	replace rse=.
	replace rse=`semean'/`mean_out' if `mean_out' > 0.0 & !missing(`mean_out')

	/* CALCULATE THE RELATIVE STANDARD ERROR OF NATURAL LOG P */
	replace rselnp=.
	replace rselnp=rse/(abs(log(`mean_out'))) if `mean_out' <= 0.5 & `mean_out' > 0.0
	replace rselnp=rse*(`mean_out'/(1-`mean_out'))/(abs(log(1-`mean_out'))) if `mean_out' < 1.0 & `mean_out' > 0.5
	
	/*CALCULATE THE EFFECTIVE SAMPLE SIZE*/
	local deffmean = (`nsum' * (`semean'*`mean_out')^2) / (`mean_out'*(1-`mean_out'))
	local effnsum=`nsum'/`deffmean'
	
	/*SUPRESSION RULE FOR PREVALENCE RATES*/
	

	replace suprule1a = 0
	replace suprule1b = 0
	replace suprule1c = 0
	replace suprule2 = 0
	replace suprule3 = 0
	
	replace suprule1a=1 if rselnp > 0.175 & !missing(rselnp)
	replace suprule1b=1 if `mean_out' < .00005 
	replace suprule1c=1 if `mean_out' >= .99995
	replace suprule2=1 if `effnsum' < 68 & !missing(`nsum')
	replace suprule3=1 if `nsum' < 100 & !missing(`nsum')
	
	global suppress=0
	if suprule1a == 1 | suprule1b == 1 | suprule1c ==1 | suprule2==1 | suprule3==1 {
		global suppress = 1
	}
end

/*********************
*Get sample size of the groups
*********************/
foreach x in ado ya {
	
	tab `x'group, matcell(freq)
	mat samplesize = freq
	gl `x'_n1 = samplesize[1,1]
	gl `x'_n2 = samplesize[2,1]
	gl `x'_n3 = samplesize[3,1]
	gl `x'_n4 = samplesize[4,1]
	gl `x'_n5 = samplesize[5,1]
}



global varlist mrjyr cocyr methamyr trqanyyr bngdrkmon ///
				aud OUD mrjud cocud mthud trqud anyothersud ///
				stndscig stndsmj stndalc stnddnk difobtmrj difobtcoc difobtcrk difobther appdrgmon2 ///
				yflpkcg2 yfltmrj2 yflmjmo yfladly2 frdpcig2 frdmevr2 frdmjmon frdadly2 ///
				prpkcig2 prmjevr2 prmjmo praldly2 ///
				rskcigpkd rskmrjmon rskmrjwk rsklsdtry rsklsdwk rskhertry rskherwk rskcocmon rskcocwk rskbngdly rskbngwk ///
				
				
global varlist	yetlknon yetlkpar yetlkbgf yetlkota yetlksop scholfl yelstgrd ///
				yevioprv yedgprgp yeslfhlp yeprgstd ///
				parchkhw parhlphw prchore2 prlmttv2 parlmtsn prgdjob2 prproud2 ///
				yofight2 yohgun2 yosell2 yostole2 yoattak2
				


*Set column globals to print
foreach gr in ado ya {

	*Globals for Column Numbers, depends on which group
	if "`gr'" == "ado" {
		*Print percent and CI
		gl `gr'_colp1 = "B"
		gl `gr'_colp2 = "C"
		gl `gr'_colp3 = "E"
		gl `gr'_colp4 = "G"
		gl `gr'_colp5 = "I"

		*Print p-value
		gl `gr'_colpv2 = "D"
		gl `gr'_colpv3 = "F"
		gl `gr'_colpv4 = "H"
		gl `gr'_colpv5 = "J"
	}
	else if "`gr'" == "ya" {
		*Print percent and CI
		gl `gr'_colp1 = "K"
		gl `gr'_colp2 = "M"
		gl `gr'_colp3 = "P"
		gl `gr'_colp4 = "S"
		gl `gr'_colp5 = "U"

		*Print p-value
		gl `gr'_colpv2 = "O"
		gl `gr'_colpv3 = "R"
		gl `gr'_colpv4 = "T"
		gl `gr'_colpv5 = "W"
	}		
}		
		
local row = 6
*local row = 58
		
*Loop 1: through each variables
foreach var in $varlist {

		di "`var'"
		
		*For most of these variables we only want the second level, but we tab and take all the levels for a select few
		if inlist("`var'", "scholfl", "yelstgrd") {
			*Save number of values for this variable
			tab `var'
			gl max = r(r)  //number of row/values of the var
			gl min = 1
		}
		else {
			gl min = 2
			gl max = 2
		}
		
		di "Min = $min , Max = $max"
		
		*** Calculate proportion in each group ***
		svy: prop `var', over(adogroup) 
		global df_ado = e(df_r)
		di $df_ado
		estimates store myprop_ado
		
		/* only run for adolescents, don't run below for YA
		* ADOONLY: Don't run for young adults for the variables that are only asked to adolescents
		if !inlist("`var'", "parentshh", "faithact", "schact") {
			svy: prop `var', over(yagroup) 
			global df_ya = e(df_r)
			di $df_ya
			estimates store myprop_ya
		}
		*/

		*Loop 2: through levels of row variable
		forvalues i=${min}/${max} {
			
			di "i = `i'"
			
			*Loop 3: through adolescent/young adult groups
			*ADOONLY: For certain vars, only run for adolescents because the variable is only asked to adolescents
			/* EDIT: FOR NOW, RUN ONLY FOR ADOLESCENTS FOR ALL VARIABLES
			if inlist("`var'", "") {
				gl grouplist ado
			}
			else {
				gl grouplist ado ya
			}
			
			gl grouplist ado
			*/
			
			foreach gr in $grouplist {
			
				estimates restore myprop_`gr'

				*Save proportion, se,and CIs for each SUD group
				*Call estimates as _prop_`i':_subpop_`j' , where prop_`i' is the level of the variable and _subpop_`j' is the SUD group
				forvalues j=1/5{
				
					gl prop_`i'_`j' = _b[_prop_`i':_subpop_`j']
					gl se_`i'_`j' = _se[_prop_`i':_subpop_`j']
					gl cil_`i'_`j': display %3.1f round((${prop_`i'_`j'} - invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)
					gl ciu_`i'_`j': display %3.1f round((${prop_`i'_`j'} + invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)

				}
				
				*Difference between all groups and group 1
				nlcom (dif2: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_2) (dif3: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_3) (dif4: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_4) (dif5: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_5), post
				
				*Save p-values
				global pval2 = 2*ttail(${df_`gr'},abs(_b[dif2]/_se[dif2]))
				global pval3 = 2*ttail(${df_`gr'},abs(_b[dif3]/_se[dif3]))
				global pval4 = 2*ttail(${df_`gr'},abs(_b[dif4]/_se[dif4]))
				global pval5 = 2*ttail(${df_`gr'},abs(_b[dif5]/_se[dif5]))
				
				*Run p-value through star program
				stars $pval2 2
				stars $pval3 3
				stars $pval4 4
				stars $pval5 5
						
				*Call suppression program to see if we need to suppress. Cycle through OUD groups
				forvalues j=1/5{
					
					*Call suppression program. Have to call the proportion, se, and sample size (n), saved above this loop
					di "Suppress OUD group `j'"
					suppressprog ${prop_`i'_`j'} ${se_`i'_`j'} ${`gr'_n`j'}
					
					*Format percent to print
					gl printmean: display %4.1f ${prop_`i'_`j'}*100
					di "${printmean} (${cil_`i'_`j'}, ${ciu_`i'_`j'})"
					
					if "$suppress" == "0" {
						putexcel ${`gr'_colp`j'}`row' = "${printmean} (${cil_`i'_`j'}, ${ciu_`i'_`j'})"
						if "`j'" != "1" {
							putexcel ${`gr'_colpv`j'}`row' = "${star`j'}"
						}
					}
					else if "$suppress" == "1" {
						*Put "n.a." instead of proportion and CI
						putexcel ${`gr'_colp`j'}`row' = "n.a."
						
						*If not the first group and the estimate is suppressed, write over the p-value
						if "`j'" != "1" {
							putexcel ${`gr'_colpv`j'}`row' = "n.a."
						}
					}
				} //end suppression program and printing
			
			} //end gr loop of adolescent/YA group
			
			*Print variable name to check
			putexcel X`row' = "`var' = `i'"	
			
			*Go to the next row for the next level of this variable
			local ++row
			sleep 500
			
		} //end i loop, through levels of variable
		
		*Only skip a row after certain variables
		if inlist("`var'", "bngdrkmon", "anyothersud", "stnddnk", "appdrgmon2", "yfladly2", "frdadly2", "praldly2", "rskbngwk", "yetlksop") ///
			| inlist("`var'", "scholfl", "yelstgrd", "yeprgstd", "prproud2", "yoattak2", "sexident") {
			local ++row
			sleep 500
		}
		
} // end var loop of variables
	

	
/*********************************************************************************************	
*********************************************************************************************
*********************************************************************************************
Run Table 2 on variables that are asked to YAs
*********************************************************************************************
*********************************************************************************************
*********************************************************************************************
*********************************************************************************************/
	

global yavarlist mrjyr cocyr methamyr trqanyyr bngdrkmon ///
				aud OUD mrjud cocud mthud trqud anyothersud ///
				difobtmrj difobtcoc difobtcrk difobther appdrgmon2 ///
				rskcigpkd rskmrjmon rskmrjwk rsklsdtry rsklsdwk rskhertry rskherwk rskcocmon rskcocwk rskbngdly rskbngwk ///
				



*Set column globals to print
foreach gr in ado ya {

	*Globals for Column Numbers, depends on which group
	if "`gr'" == "ado" {
		*Print percent and CI
		gl `gr'_colp1 = "B"
		gl `gr'_colp2 = "C"
		gl `gr'_colp3 = "E"
		gl `gr'_colp4 = "G"
		gl `gr'_colp5 = "I"

		*Print p-value
		gl `gr'_colpv2 = "D"
		gl `gr'_colpv3 = "F"
		gl `gr'_colpv4 = "H"
		gl `gr'_colpv5 = "J"
	}
	else if "`gr'" == "ya" {
		*Print percent and CI
		gl `gr'_colp1 = "K"
		gl `gr'_colp2 = "M"
		gl `gr'_colp3 = "P"
		gl `gr'_colp4 = "S"
		gl `gr'_colp5 = "U"

		*Print p-value
		gl `gr'_colpv2 = "O"
		gl `gr'_colpv3 = "R"
		gl `gr'_colpv4 = "T"
		gl `gr'_colpv5 = "W"
	}		
}		
		
local row = 6
*local row = 58
		
*Loop 1: through each variables
foreach var in $yavarlist {

		di "`var'"
		
		*For most of these variables we only want the second level, but we tab and take all the levels for a select few
		if inlist("`var'", "scholfl", "yelstgrd") {
			*Save number of values for this variable
			tab `var'
			gl max = r(r)  //number of row/values of the var
			gl min = 1
		}
		else {
			gl min = 2
			gl max = 2
		}
		
		di "Min = $min , Max = $max"
		
		*** Calculate proportion in each group ***
		svy: prop `var', over(yagroup) 
		global df_ya = e(df_r)
		di $df_ya
		estimates store myprop_ya
		

		*Loop 2: through levels of row variable
		forvalues i=${min}/${max} {
			
			di "i = `i'"
			
			
			foreach gr in ya {
			
				estimates restore myprop_`gr'

				*Save proportion, se,and CIs for each SUD group
				*Call estimates as _prop_`i':_subpop_`j' , where prop_`i' is the level of the variable and _subpop_`j' is the SUD group
				forvalues j=1/5{
				
					gl prop_`i'_`j' = _b[_prop_`i':_subpop_`j']
					gl se_`i'_`j' = _se[_prop_`i':_subpop_`j']
					gl cil_`i'_`j': display %3.1f round((${prop_`i'_`j'} - invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)
					gl ciu_`i'_`j': display %3.1f round((${prop_`i'_`j'} + invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)

				}
				
				*Difference between all groups and group 1
				nlcom (dif2: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_2) (dif3: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_3) (dif4: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_4) (dif5: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_5), post
				
				*Save p-values
				global pval2 = 2*ttail(${df_`gr'},abs(_b[dif2]/_se[dif2]))
				global pval3 = 2*ttail(${df_`gr'},abs(_b[dif3]/_se[dif3]))
				global pval4 = 2*ttail(${df_`gr'},abs(_b[dif4]/_se[dif4]))
				global pval5 = 2*ttail(${df_`gr'},abs(_b[dif5]/_se[dif5]))
				
				*Run p-value through star program
				stars $pval2 2
				stars $pval3 3
				stars $pval4 4
				stars $pval5 5
						
				*Call suppression program to see if we need to suppress. Cycle through OUD groups
				forvalues j=1/5{
					
					*Call suppression program. Have to call the proportion, se, and sample size (n), saved above this loop
					di "Suppress OUD group `j'"
					suppressprog ${prop_`i'_`j'} ${se_`i'_`j'} ${`gr'_n`j'}
					
					*Format percent to print
					gl printmean: display %4.1f ${prop_`i'_`j'}*100
					di "${printmean} (${cil_`i'_`j'}, ${ciu_`i'_`j'})"
					
					if "$suppress" == "0" {
						putexcel ${`gr'_colp`j'}`row' = "${printmean} (${cil_`i'_`j'}, ${ciu_`i'_`j'})"
						if "`j'" != "1" {
							putexcel ${`gr'_colpv`j'}`row' = "${star`j'}"
						}
					}
					else if "$suppress" == "1" {
						*Put "n.a." instead of proportion and CI
						putexcel ${`gr'_colp`j'}`row' = "n.a."
						
						*If not the first group and the estimate is suppressed, write over the p-value
						if "`j'" != "1" {
							putexcel ${`gr'_colpv`j'}`row' = "n.a."
						}
					}
				} //end suppression program and printing
			
			} //end gr loop of adolescent/YA group
			
			*Print variable name to check
			putexcel X`row' = "`var' = `i'"	
			
			*Go to the next row for the next level of this variable
			local ++row
			sleep 500
			
		} //end i loop, through levels of variable
		
		*Only skip a row after certain variables
		if inlist("`var'", "bngdrkmon", "anyothersud", "stnddnk", "appdrgmon2", "yfladly2", "frdadly2", "praldly2", "rskbngwk", "yetlksop") ///
			| inlist("`var'", "scholfl", "yelstgrd", "yeprgstd", "prproud2", "yoattak2", "sexident") {
			local ++row
			sleep 500
		}
		
} // end var loop of variables
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
/* ONLY RUN FOR ADOLESCENTS


/*************************************************************************
Look through comp groups, comparing adolescent and young adult groups
*************************************************************************/
*Save column values to print stars for differences between ado/ya groups
	gl comp_colpv1 = "L"
	gl comp_colpv2 = "N"
	gl comp_colpv3 = "Q"
	gl comp_colpv5 = "V"



local row = 6
		
*Loop 1: through each variables
foreach var in $varlist {
		
		*For most of these variables we only want the second level, but we tab and take all the levels for a select few
		if inlist("`var'", "") {
			*Save number of values for this variable
			tab `var'
			gl max = r(r)  //number of row/values of the var
			gl min = 1
		}
		else {
			gl min = 2
			gl max = 2
		}

		di "Min = $min , Max = $max"
		

		*** Calculate proportion in each group ***
		*Can run on all variables, even the ones we don't compare, since as long as there's data for one group we don't get an error
		svy: prop `var', over(comp1) 
		global df1 = e(df_r)
		di $df1
		estimates store myprop_1
		
		svy: prop `var', over(comp2) 
		global df2 = e(df_r)
		di $df2
		estimates store myprop_2

		svy: prop `var', over(comp3) 
		global df3 = e(df_r)
		di $df3
		estimates store myprop_3
		
		svy: prop `var', over(comp5) 
		global df5 = e(df_r)
		di $df5
		estimates store myprop_5
		
		*Loop 2: through levels of row variable
		forvalues i=${min}/${max} {
			
			di "i = `i'"
			
			*Loop 3: through adolescent/young adult comparison groups, but only for variables that have data for both groups
			*ADOONLY: add in variables that don't have data for both groups
			if !inlist("`var'", "") {
				
				foreach gr in 1 2 3 5 {

					di "Variable: `var', Group `gr'"
					estimates restore myprop_`gr'

					*Save proportion, se,and CIs for each SUD group
					*Call estimates as _prop_`i':_subpop_`j' , where prop_`i' is the level of the variable and _subpop_`j' is the SUD group
					forvalues j=1/2{
					
						gl prop_`i'_`j' = _b[_prop_`i':_subpop_`j']
						gl se_`i'_`j' = _se[_prop_`i':_subpop_`j']
						gl cil_`i'_`j': display %3.1f round((${prop_`i'_`j'} - invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)
						gl ciu_`i'_`j': display %3.1f round((${prop_`i'_`j'} + invttail(e(df_r),0.025)*${se_`i'_`j'})*100, .1)

					}
					
					*Difference between group 1 (adolescents) and group 2 (young adults)
					nlcom (dif2: [_prop_`i']_subpop_1 - [_prop_`i']_subpop_2), post
					
					*Save p-values
					global pval = 2*ttail(${df`gr'},abs(_b[dif2]/_se[dif2]))
					
					*Run p-value through star program
					plus $pval 1
							
					*Call suppression program for group 2
					*Have to call the proportion, se, and sample size (n), saved above this loop
					suppressprog ${prop_`i'_2} ${se_`i'_2} ${ya_n`gr'}
							
						if "$suppress" == "0" {
							putexcel ${comp_colpv`gr'}`row' = "${plus1}"
						}
						else if "$suppress" == "1" {
							*Put "n.a." instead of proportion and CI
							putexcel ${comp_colpv`gr'}`row' = "n.a."
							
					} 
				
				} //end gr loop of adolescent/YA group
				
			} // end if statement that ignores certain variables
			
			*Print variable name to check
			putexcel Y`row' = "`var' = `i'"	
			
			*Go to the next row for the next level of this variable
			local ++row
			sleep 500
			
		} //end i loop, through levels of variable
		
			*Only skip a row after certain variables
		if inlist("`var'", "bngdrkmon", "anyothersud", "stnddnk", "appdrgmon2", "yfladly2", "frdadly2", "praldly2", "rskbngwk") {
			local ++row
			sleep 500
		}
			
} // end var loop of variables
	
	

*/

