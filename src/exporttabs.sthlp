{smcl}
{* *! version 1.0.0 22JAN2026}{...}
{hline}
{vieweralsosee "" "--"}{...}
{vieweralsosee "tabulate" "help tabulate"}{...}
{vieweralsosee "putexcel" "help putexcel"}{...}
{viewerjumpto "Syntax" "exporttabs##syntax"}{...}
{viewerjumpto "Description" "exporttabs##description"}{...}
{viewerjumpto "Options" "exporttabs##options"}{...}
{viewerjumpto "Examples" "exporttabs##examples"}{...}
{viewerjumpto "Remarks" "exporttabs##remarks"}{...}
{viewerjumpto "Stored results" "exporttabs##results"}{...}
{viewerjumpto "References" "exporttabs##references"}{...}
{viewerjumpto "Author" "exporttabs##author"}{...}
{title:Title}

{phang}
{bf:exporttabs} {hline 2} Export frequency distributions and cross-tabulations to Excel

{marker syntax}{...}
{title:Syntax}

{p 8 16 2}
{cmd:exporttabs} [{varlist}] [{cmd:if} {it:exp}] [{cmd:in} {it:range}] {cmd:using} {it:{help filename}} [{cmd:,} {it:options}]

{synoptset 26 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:Main}
{synopt:{opt by(varlist)}}variables for cross-tabulation{p_end}
{synopt:{opt tabopt(string)}}display options: row, col, or cell percentages{p_end}
{synopt:{opt sheet(name)}}Excel worksheet name{p_end}

{syntab:Advanced}
{synopt:{opt missing}}include missing values in tables{p_end}
{synopt:{opt nolabel}}display raw values instead of value labels{p_end}
{synopt:{opt force}}overwrite existing Excel file without prompt{p_end}
{synopt:{opt replace}}replace sheet if exists (default: append){p_end}
{synoptline}
{p2colreset}{...}

{marker description}{...}
{title:Description}

{pstd}
{cmd:exporttabs} exports one-way frequency tables and two-way cross-tabulations 
from Stata to Microsoft Excel (.xlsx) format. The command creates professionally 
formatted tables suitable for reports, presentations, and publications.

{pstd}
When no {opt by()} option is specified, {cmd:exporttabs} produces one-way 
frequency tables showing counts and percentages for each variable. When 
{opt by()} is specified, two-way cross-tabulations are created for each 
combination of variables.

{pstd}
The command automatically handles:
{p 8 12}• Variable and value labels{p_end}
{p 8 12}• Numeric and string variables{p_end}
{p 8 12}• Missing value treatment{p_end}
{p 8 12}• Percentage calculations{p_end}
{p 8 12}• Excel formatting (bold headers, number formats){p_end}

{pstd}
Output includes table titles, row/column headers, frequency counts, 
percentages (when requested), and marginal totals. The Excel file is 
structured with clear separation between tables for easy navigation.

{marker options}{...}
{title:Options}

{dlgtab:Main}

{phang}
{opt by(varlist)} specifies one or more variables to use for cross-tabulation.
For each variable in {it:varlist}, {cmd:exporttabs} creates a separate 
two-way table with each variable in {opt by()}. If a variable appears in 
both {it:varlist} and {opt by()}, that combination is skipped. When 
{opt by()} is not specified, one-way frequency tables are produced.

{pmore}
Example: {cmd:exporttabs gender education using "results.xlsx", by(region)}
creates two tables: gender × region and education × region.

{phang}
{opt tabopt(string)} specifies what values to display in table cells. 
Options are:

{pmore}
{opt row}: Display row percentages (percentage of row total). Useful when 
you want to compare distributions across columns for each row category.

{pmore}
{opt col}: Display column percentages (percentage of column total). Useful 
when you want to compare distributions across rows for each column category.

{pmore}
{opt cell}: Display cell percentages (percentage of grand total). Shows 
each cell's contribution to the overall total.

{pmore}
If {opt tabopt()} is not specified, cells display frequency counts. 
Marginal totals always show frequency counts regardless of this option.

{phang}
{opt sheet(name)} specifies the name of the Excel worksheet where tables 
will be written. The default is "Tables". Worksheet names:
{p 8 12}• Cannot exceed 31 characters{p_end}
{p 8 12}• Cannot contain: : \ / ? * [ ] {p_end}
{p 8 12}• Cannot begin or end with an apostrophe (''){p_end}
{p 8 12}• Cannot be blank{p_end}

{dlgtab:Advanced}

{phang}
{opt missing} includes missing values in the tabulations. By default, 
missing values (system missing ., extended missing .a-.z for numeric 
variables, and empty strings "" for string variables) are excluded from 
tables. When {opt missing} is specified, these values are treated as 
valid categories.

{phang}
{opt nolabel} displays raw numeric values instead of value labels. This 
is useful when:
{p 8 12}• Value labels are very long{p_end}
{p 8 12}• You need the numeric codes for further analysis{p_end}
{p 8 12}• The dataset has no value labels defined{p_end}

{phang}
{opt force} overwrites an existing Excel file without prompting for 
confirmation. Use with caution as deleted files cannot be recovered. 
Without {opt force}, Stata will display an error if the file already exists.

{phang}
{opt replace} replaces the specified worksheet if it already exists in 
the Excel file. By default, {cmd:exporttabs} appends to existing worksheets. 
This option requires {opt force} if the entire file needs to be overwritten.

{marker examples}{...}
{title:Examples}

{pstd}
Setup: Load example dataset

{phang2}{cmd:. sysuse auto, clear}{p_end}

{pstd}
Example 1: Basic one-way frequency tables for all variables

{phang2}{cmd:. exporttabs using "auto_frequencies.xlsx"}{p_end}

{pstd}
Example 2: One-way tables for specific variables

{phang2}{cmd:. exporttabs price mpg rep78 using "selected_vars.xlsx"}{p_end}

{pstd}
Example 3: Two-way cross-tabulation

{phang2}{cmd:. exporttabs foreign using "crosstab.xlsx", by(rep78)}{p_end}

{pstd}
Example 4: Cross-tabulation with row percentages

{phang2}{cmd:. exporttabs foreign using "crosstab_row.xlsx", by(rep78) tabopt(row)}{p_end}

{pstd}
Example 5: Multiple cross-tabulations

{phang2}{cmd:. exporttabs price mpg using "multiple_crosstabs.xlsx", by(foreign rep78)}{p_end}
{pmore}Creates four tables: price×foreign, price×rep78, mpg×foreign, mpg×rep78{p_end}

{pstd}
Example 6: With if condition

{phang2}{cmd:. exporttabs price mpg using "domestic.xlsx" if foreign == 0}{p_end}

{pstd}
Example 7: Include missing values

{phang2}{cmd:. exporttabs rep78 using "with_missing.xlsx", missing}{p_end}

{pstd}
Example 8: Custom worksheet name

{phang2}{cmd:. exporttabs using "analysis.xlsx", sheet("Auto Industry Data")}{p_end}

{pstd}
Example 9: Complex combination

{phang2}{cmd:. exporttabs price mpg weight using "full_report.xlsx", by(foreign) tabopt(col) if price > 5000}{p_end}

{pstd}
Example 10: Using variable ranges

{phang2}{cmd:. exporttabs price-weight using "range_vars.xlsx", by(foreign)}{p_end}

{marker remarks}{...}
{title:Remarks}

{pstd}
{bf:Variable and value labels}

{pstd}
{cmd:exporttabs} automatically uses variable labels and value labels when 
they are defined. Variable labels appear in table titles, and value labels 
appear as category names. To check or define labels:

{phang2}{cmd:. describe}{p_end}
{phang2}{cmd:. label list}{p_end}
{phang2}{cmd:. label variable varname "New variable label"}{p_end}
{phang2}{cmd:. label define yesno 1 "Yes" 0 "No"}{p_end}
{phang2}{cmd:. label values varname yesno}{p_end}

{pstd}
{bf:Excel formatting}

{pstd}
The command applies the following Excel formatting:
{p 8 12}• Bold font for headers and totals{p_end}
{p 8 12}• Two decimal places for percentages{p_end}
{p 8 12}• Proper number formatting for frequencies{p_end}
{p 8 12}• Clear separation between tables (blank rows){p_end}

{pstd}
For additional formatting (colors, borders, alignment), edit the Excel 
file manually after export.

{pstd}
{bf:Performance considerations}

{pstd}
For large datasets or many variables, consider:
{p 8 12}• Using {cmd:if} or {cmd:in} to limit observations{p_end}
{p 8 12}• Processing variables in batches{p_end}
{p 8 12}• Closing Excel before running the command{p_end}
{p 8 12}• Using local drives instead of network drives{p_end}

{pstd}
{bf:Limitations}

{pstd}
{cmd:exporttabs} has the following limitations:
{p 8 12}• Maximum 1,048,576 rows per worksheet (Excel limit){p_end}
{p 8 12}• Maximum 16,384 columns per worksheet (Excel limit){p_end}
{p 8 12}• Three-way or higher cross-tabulations not supported{p_end}
{p 8 12}• Weighted frequencies not supported{p_end}
{p 8 12}• Statistical tests (chi-square, etc.) not included{p_end}

{marker results}{...}
{title:Stored results}

{pstd}
{cmd:exporttabs} does not store results in Stata's memory. All output 
is written directly to the specified Excel file. During execution, the 
command displays:

{pmore}
Progress indicators: Shows each table being processed{p_end}
{pmore}
Completion summary: Number of tables created and file location{p_end}
{pmore}
Usage tips: Helpful reminders about command options{p_end}

{pstd}
To capture the output programmatically, use the {cmd:quietly} prefix:

{phang2}{cmd:. quietly exporttabs using "results.xlsx"}{p_end}

{marker error_messages}{...}
{title:Error messages}

{pstd}
Common error messages and solutions:

{pmore}
{err:file filename.xlsx already exists}{p_end}
{pmore2}Solution: Use {opt force} option or choose different filename{p_end}

{pmore}
{err:invalid sheet name}{p_end}
{pmore2}Solution: Sheet name must be ≤31 characters, no special characters{p_end}

{pmore}
{err:insufficient observations}{p_end}
{pmore2}Solution: Check {cmd:if/in} conditions or use {opt missing} option{p_end}

{pmore}
{err:Excel file could not be opened}{p_end}
{pmore2}Solution: Close Excel, check file permissions, ensure disk space{p_end}

{pmore}
{err:variable not found}{p_end}
{pmore2}Solution: Check variable names, use {cmd:describe} to list variables{p_end}

{pmore}
{err:too many rows/columns}{p_end}
{pmore2}Solution: Reduce number of variables or categories, use {cmd:if/in}{p_end}

{marker references}{...}
{title:References}

{pstd}
StataCorp. 2023. Stata: Release 18. Statistical Software. College Station, TX: StataCorp LLC.

{pstd}
Microsoft Corporation. 2023. Excel specifications and limits. 
{browse "https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3"}

{marker author}{...}
{title:Author}

{pstd}
Md. Redoan Hossain Bhuiyan{p_end}
{pstd}
Email: redoanhossain630@gmail.com{p_end}
{pstd}
Website: {browse "https://www.example.com/stata-tools"}{p_end}

{marker acknowledgment}{...}
{title:Acknowledgment}

{pstd}
The development of {cmd:exporttabs} was inspired by user requests for 
a simple, automated way to export Stata tables to Excel. Special thanks 
to the Stata user community for feedback and testing.

{title:Also see}

{psee}
Manual: {manhelp tabulate R}, {manhelp putexcel P}, {manhelp export_excel P}

{psee}
Online: {browse "http://www.stata.com/support/faqs/"}{p_end}

{hline}
