*******************************************************************************
*   Version 0.2: 2023-02-27
*******************************************************************************
*   Dennis Föste-Eggers	
*   
*   German Centre for Higher Education Research and Science Studies (DZHW)
*   Lange Laube 12, 30159 Hannover         
*   Phone: +49-(0)511 450 670-114	
*   
*   E-Mail (1): dennis.foeste@gmail.com 		
*   E-Mail (2): dennis.foeste@outlook.de
*   E-Mail (3): dennis.foeste@gmx.de
*   
*   E-Mail (4): foeste-eggers@dzhw.eu
*******************************************************************************
*   Program name: str_harm.ado     
*   Program purpose: An easy to use harmonization of strings.			
*******************************************************************************
*   Changes made:
*   Version 0.1: added GPL 
*   Version 0.2: minimal changes of syntax
*******************************************************************************
*   License: GPL (>= 3)
*     
*   str_harm.ado for Stata
*   Copyright (C) 2023 Foeste-Eggers, Dennis 
*   
*   This program is free software: you can redistribute it and/or modify
*   it under the terms of the GNU General Public License as published by
*   the Free Software Foundation, either version 3 of the License, or
*   (at your option) any later version.
*   
*   This program is distributed in the hope that it will be useful,
*   but WITHOUT ANY WARRANTY; without even the implied warranty of
*   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
*   GNU General Public License for more details.
*   
*   You should have received a copy of the GNU General Public License
*   along with this program.  If not, see <https://www.gnu.org/licenses/>.
*   
*******************************************************************************
*   Citation: This code is © D. Foeste-Eggers, 2023, and it is made 
*				 available under the GPL license enclosed with the software.
*
*!			Over and above the legal restrictions imposed by this license, if !
*!          you use this program for any (academic) publication then you are  !
*! 			obliged to provide proper attribution.                            !
*
*   D. Foeste-Eggers str_harm.ado for Stata, v0.2 (2023). 
*           [weblink].
*
*******************************************************************************
cap program drop str_harm 
program define str_harm  , nclass
    version 15
	
    syntax varlist [if] [in] [, 			///
                                GENerate(namelist)	    	 /// namelist
                                REPLACE 	            	 /// replace vars
                                SUFfix(namelist max=1)  	 /// undocumented -tbd
                                PREfix(namelist max=1)  	 /// undocumented -tbd
                                MISSINGS(passthru) 			 /// undocumented -tbd
                                MISETs(passthru)        	 /// undocumented -tbd
								viewreplacements			 /// 
                                xlsx_replacements(passthru)	 /// should be synonyms possible as well  / 
                                sheet_replacements(passthru) /// undocumented -tbd
								LOWer 						 /// lower case instead of upper 
                                cellrange(passthru)]	 	 // undocumented -tbd  
                                           

								
    di ""
    di "Harmonization of strings via str_harm.ado by Foeste-Eggers (2023,V0.2)"
    di ""
    if `"`generate'"' ~= `""' & `"`replace'"' ~= `""' ///
      & `"`suffix'"'  == `""' & `"`prefix'"'  == `""' {
          
        di as error "replace option was ignored due to use of generate option"
        local replace = ""
    }
    if `"`generate'"'  ~= `""'  ///
      & (`"`suffix'"'  ~= `""' | `"`prefix'"' ~= `""') {
          
        di as error "generate option was ignored due to use pre- or suffixes"
        local generate = ""
    }
    if `"`generate'"' == `""' & `"`replace'"' == `""' ///
      & `"`suffix'"'  == `""' & `"`prefix'"'  == `""' {
          
        di as error "neither replace option nor generate option were used"
        error 198
    }
	
	qui if `"`xlsx_replacements'"' ~= `""' {
		local xlsx_replacements = subinstr(`"`xlsx_replacements'"',`"xlsx_replacements("'," ",1)
		local xlsx_replacements = substr(`"`xlsx_replacements'"',1,`=length(`"`xlsx_replacements'"')-1')
		local xlsx_replacements = trim(`"`xlsx_replacements'"')
        *noi di `"`xlsx_replacements' `sheet_replacements'"'
		if `"`sheet_replacements'"' == `""' local sheet_replacements = `"Sheet1"'
		else {
			local sheet_replacements = subinstr(`"`sheet_replacements'"',`"sheet_replacements("',"",1)
			local sheet_replacements = substr(`"`sheet_replacements'"',1,`=length(`"`sheet_replacements'"')-1')
			local sheet_replacements = trim(`"`sheet_replacements'"')
		}
		
		* preserve
		tempfile dim 
		if `"`dim'"'~=`""' save `dim' , replace
		
		
			cap noi import excel using `xlsx_replacements', describe //  clear allstring cellrange(A:C)  
            noi return list
            forvalues ws = 1(1)`=scalar(r(N_worksheet))' {
                    if `"`r(worksheet_`ws')'"' == `"`sheet_replacements'"' local ws_is = `ws'
            }
            if `"`ws_is'"'=="" {
                          forvalues ws = 1(1)`=scalar(r(N_worksheet))' {
                                if `"`r(worksheet_`ws')'"' == `"Tabelle1"' local ws_is = `ws'
                                if `"`r(worksheet_`ws')'"' == `"Tabelle1"' local sheet_replacements = `"Tabelle1"'
                          }
            }
                
            if  `"`ws_is'"'=="" noi di as error `"Worksheet not found"'
            else {
                noi di `"cap import excel using `xlsx_replacements', sheet(`sheet_replacements') clear allstring cellrange(`r(range_`ws_is')')"'
                cap import excel using `xlsx_replacements', sheet(`"`sheet_replacements'"') clear allstring cellrange(`r(range_`ws_is')')
                if _rc noi di as error `"`xlsx_replacements' could not be used"'
                noi di `"Cellrange is `r(range_`ws_is')'"'
            }
            
            local ws_is = "" 
            * noi return list
			* if _rc cap import excel using `"`xlsx_replacements'"', sheet(`"Sheet1"') 		clear allstring 
			* if _rc cap import excel using `"`xlsx_replacements'"', sheet(`"Tabelle1"') 		clear allstring 
			
			noi ds 
			local var_list = `"`r(varlist)'"' 
			*keep A B 
			local var_list = trim(subinstr(`" `var_list' "' ',`" A "'," ",.))
			local var_list = trim(subinstr(`" `var_list' "' ',`" B "'," ",.))
			local var_list = trim(subinstr(`" `var_list' "' ',`" C "'," ",.))
			cap drop `var_list'
			
			noi ds 
			local var_list = `"`r(varlist)'"' 
			* if strpos(`"`var_list'"',"C")>0 {
			* 	keep A B C // to force error if A or B is missing
			* 	gen syntax = `"cap replace "' + "\`" + "str_harm" + "'" + ///
			* 	`"= usubinstr("' + "\`" + "str_harm" + "'" + `",""' + A + `"",""' + B + `"",.) "' ///
			* 	+ C 
			* }
			* else {
			* 	keep A B  // to force error if A or B is missing
			replace A = "\`" + `"""' + A + `"""' + "'"
			replace B = "\`" + `"""' + B + `"""' + "'"
			gen syntax = `"cap replace "' + "\`" + "str_harm" + "'" + ///
			`"= usubinstr("' + "\`" + "str_harm" + "'" + `","' + upper(A) + `","' + upper(B) + `",.) "' // 
            cap replace syntax = syntax + " " + C
			* }
            noi di as result `"resulting syntax (from `xlsx_replacements')"'
			noi list syntax, noheader
			
			keep syntax*
			
			tempfile syntax_sh
			
			if `"`syntax_sh'"'~=`""' outfile syntax using `syntax_sh'.do, noquote replace wide
	
			if `"`viewreplacements'"'~=`""' view `syntax_sh'.do

    

			
		* restore
		use `dim', clear
		
	}
			
			
    	
	tempvar touse 
	mark `touse' `if' `in'
	
	local n = 0
	qui foreach var of varlist `varlist' {
	    local ++n 
		local undo = 0
		* noi sum _all
		tempvar str_harm blanks
        cap drop `str_harm'
	    clonevar `str_harm' = `var' if `touse'
        
        replace `str_harm' = usubinstr(`str_harm',`"`=char(9)'"'," ",.)
        replace `str_harm' = usubinstr(`str_harm',`"`=char(10)'"'," ",.)
        replace `str_harm' = usubinstr(`str_harm',`"`=char(13)'"'," ",.)
		
		
        
        replace `str_harm' = upper(strtrim(`str_harm'))
        
        replace `str_harm' = usubinstr(`str_harm',"Ä","AE", . )
        replace `str_harm' = usubinstr(`str_harm',"Ü","UE", . )
        replace `str_harm' = usubinstr(`str_harm',"Ö","OE", . )
        replace `str_harm' = usubinstr(`str_harm',"ß","SS", . )
        replace `str_harm' = usubinstr(`str_harm',"ä","AE", . )
        replace `str_harm' = usubinstr(`str_harm',"ü","UE", . )
        replace `str_harm' = usubinstr(`str_harm',"ö","OE", . )
        
        *local multipleblanks = 1
        *while `multipleblanks'>0 {
        *    
        *    replace `str_harm' = subinstr(`str_harm',"  "," ", . )
        *    
        *    cap drop `blanks'
        *    gen `blanks' = strpos(`str_harm',"  ") if `touse'
        *    sum `blanks'                           if `touse' , meanonly
        *    local multipleblanks = `r(max)'
        *    
        *}

        replace `str_harm' = subinstr(`str_harm',"&AACUTE;","A",.)
        replace `str_harm' = subinstr(`str_harm',"&EACUTE;","E",.)
        replace `str_harm' = subinstr(`str_harm',"&IACUTE;","I",.)
        replace `str_harm' = subinstr(`str_harm',"&OACUTE;","O",.)
        replace `str_harm' = subinstr(`str_harm',"&UACUTE;","U",.)
        replace `str_harm' = subinstr(`str_harm',"&AGRAVE;","A",.)
        replace `str_harm' = subinstr(`str_harm',"&EGRAVE;","E",.)
        replace `str_harm' = subinstr(`str_harm',"&IGRAVE;","I",.)
        replace `str_harm' = subinstr(`str_harm',"&OGRAVE;","O",.)
        replace `str_harm' = subinstr(`str_harm',"&UGRAVE;","U",.)
        replace `str_harm' = subinstr(`str_harm',"&CCEDIL;","C",.)
		
		if `"`syntax_sh'"' ~= `""' {
			include `syntax_sh'.do
		}
		
		local multipleblanks = 1
        while `multipleblanks'>0 {
            
            replace `str_harm' = subinstr(`str_harm',"  "," ", . )
            
            cap drop `blanks'
            gen `blanks' = strpos(`str_harm',"  ") if `touse'
            sum `blanks'                           if `touse' , meanonly
            local multipleblanks = `r(max)'
            
        }
		
		if `"`lower'"' ~= `""' {
			replace `str_harm' = lower(strtrim(`str_harm'))
		}
        
        if `"`generate'"' ~= `""' {
            local name : word `n' of `generate'
            rename `str_harm' `name'           
        }
        else if `"`replace'"' ~= `""' {
            replace `var' = `str_harm' if `touse'
        }

	}
    
    
end

