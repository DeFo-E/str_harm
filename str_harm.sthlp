{smcl}
{* ado file version 0.2: 2023-02-27 }{...}
help for {hi:str_harm}{right: {browse "mailto:dennis.foeste@outlook.de":Dennis Foeste-Eggers}}
{hline}

{title:Title}

{phang}
{bf:str_harm} {hline 1} Easy-to-use program for harmonizing/canonizing of string variables. 

{p 8 17 2}
{cmd:str_harm} {varlist} {ifin} [{cmd:,} {it:save_options}  {it:xlsx_input_options}]
{p_end}

{pstd}{it:save_options} You have to choose one of the three options to save the results of the string harmonization (generate, replace or suffix|prefix).  {p_end}

{pstd}{it:xlsx_input_options} You may use an Excel file as input for an intended sequence of replacements within the string variable(s). You may also specify the Excel worksheet and the cell range to be used.
    The Excel sheet must be structured as follows: column A contains the expression to be replaced, column B contains the intended replacement and column C may be used to add if/in qulifiers. 
    The rows are processed in ascending order.  {p_end}

{synoptset 35 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:{it:save_options}}
{synopt:{opth gen:erate(newvarlist)}}generates a new variable for each variable of {it:varlist} as specified in {it:newvarlist} {p_end}
{synopt:{opt replace}}each variable of {it:varlist} will be manipulated and replaced as specified{p_end}
{synopt:{opth suf:fix(namelist max=1)}}generates a new variable for each variable of {it:varlist} using the specified suffix {p_end}
{synopt:{opth pre:fix(namelist max=1)}}generates a new variable for each variable of {it:varlist} using the specified prefix {p_end}

{syntab:{it:xlsx_input_options}}
{synopt:{opth xlsx_replacements(filename)}}generates a new variable for each variable of {it:varlist} using the specified prefix {p_end}
{synopt:{opt sheet_replacements("sheetname")}}option to specify the name of Excel worksheet to be used {p_end}
{synopt:{opt cellrange:([start][:end])}}option to specify the Excel cell range to be used {p_end}
{synopt:{opt viewreplacements}}allows you to view the sequence replacements in the Stata viewer {p_end}

{synoptline}
{p2colreset}{...}


{title:Description}

{p}
{cmd:str_harm} Enables users to harmonize/canonize string variables, which might be helpful if the are to be used for merging (e. g. with data containing related context informations)
{p_end}


