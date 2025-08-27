{smcl}
{* *! version 1.0  25aug2025}{...}
{title:Title}

{phang}
{bf:exporttabs} — Export one-way and two-way tabulations to Excel (appliable for single choice question)

{title:Syntax}

{p 8 16 2}
{cmd:exporttabs} [{varlist}] {cmd:using} {it:filename} 
[{cmd:,} {opt by(varlist)} {opt tabopt(string)}]

{title:Description}

{pstd}
{cmd:exporttabs} automates the export of one-way and two-way tabulations 
to an Excel file. It is especially useful when you need to generate the 
same tables for many variables or for the entire dataset.  

{pstd}
By default, one-way frequency tables are exported for each variable in 
{it:varlist}. If no {it:varlist} is given, all variables in the dataset 
are tabulated.  

{pstd}
If the {opt by()} option is specified, cross-tabulations are produced 
between each variable in {it:varlist} and the variable(s) in {opt by()}.  
The {opt tabopt()} option allows you to pass standard {help tabulate} 
options, such as {cmd:row}, {cmd:col}, {cmd:cell}, {cmd:nofreq}, etc.  

{title:Options}

{phang}
{opt by(varlist)}  
    One or more variables to cross-tabulate against each variable in {it:varlist}.  

{phang}
{opt tabopt(string)}  
    Tabulation options controlling display. Examples:  
    {cmd:row}, {cmd:col}, {cmd:row nofreq}, {cmd:col nofreq}, {cmd:cell}.  

{title:Remarks}

{pstd}
- Results are exported to a new Excel file (default behavior: {cmd:replace}).  
- Percentages are rounded to 2 decimal places.  
- Tables are exported in plain format; you may manually add borders, 
shading, or adjust fonts in Excel.  
- Always verify totals when using percentages.  
- Works with both labeled and unlabeled categorical variables.  

{title:Examples}

{pstd}
Suppose we have 250 observations of students with an {bf:age_group} variable 
and a {bf:district} variable (5 districts: Dhaka, Cumilla, Chandpur, Gazipur, 
Cox's Bazar).  

{pstd}
One-way table (all variables by default):  
{cmd}
    . exporttabs using "01_Data_Tables_single.xlsx"
{txt}
{pstd}
Cross-tabulation with frequencies:  
{cmd}
    . exporttabs using "02_Data_Tables_cross_freq.xlsx", by(district)
{txt}
{pstd}
Cross-tabulation with column percentages:  
{cmd}
    . exporttabs using "03_Data_Tables_col.xlsx", by(district) tabopt("col")
{txt}
{pstd}
Column percentages without frequencies:  
{cmd}
    . exporttabs using "04_Data_Tables_col_nofreq.xlsx", by(district) tabopt("col nofreq")
{txt}
{pstd}
Row percentages:  
{cmd}
    . exporttabs using "05_Data_Tables_row.xlsx", by(district) tabopt("row")
{txt}
{pstd}
Row percentages without frequencies:  
{cmd}
    . exporttabs using "06_Data_Tables_row_nofreq.xlsx", by(district) tabopt("row nofreq")
{txt}
{pstd}
Cell percentages:  
{cmd}
    . exporttabs using "07_Data_Tables_cell.xlsx", by(district) tabopt("cell")
{txt}
{pstd}
No r-class results are returned. Output is directly written to Excel.

{title:Author}

{pstd}
Written by [Md. Redoan Hossain Bhuiyan, redoanhossain630@gmail.com], 2025.  


{title:Also see}

{psee}
{help biascheck}, {help optcounts}, {help detectoutlier}



