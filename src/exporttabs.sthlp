
{title:exporttabs - Export frequency and cross-tabulation tables to Excel}

{phang}
{bf:exporttabs} {hline 2} Export one-way frequency tables and two-way cross-tabulations 
to Microsoft Excel format (.xlsx) with professional formatting.

{marker syntax}{...}
{title:Syntax}

{p 8 17 2}
{cmd:exporttabs} [{varlist}] [{cmd:if} {it:exp}] [{cmd:in} {it:range}] 
{cmd:using} {it:{help filename}} [{cmd:,} {it:options}]

{synoptset 25 tabbed}{...}
{synopthdr}
{synoptline}
{syntab:Main}
{synopt:{opt by(varlist)}}variables for cross-tabulation (creates two-way tables){p_end}
{synopt:{opt tabopt(string)}}cell display: {opt row}, {opt col}, or {opt cell} percentages{p_end}
{synopt:{opt sheet(string)}}Excel sheet name (default: "Tables"){p_end}

{syntab:Data handling}
{synopt:{opt missing}}include missing values in tabulations{p_end}
{synopt:{opt nolabel}}display raw values instead of value labels{p_end}

{syntab:File handling}
{synopt:{opt force}}overwrite existing Excel file without confirmation{p_end}
{synopt:{opt replace}}replace sheet if it exists (default: append){p_end}
{synoptline}
{p2colreset}{...}

{marker description}{...}
{title:Description}

{pstd}
{cmd:exporttabs} creates professionally formatted Excel files containing 
frequency distributions (one-way tables) and cross-tabulations (two-way tables). 
The program automatically handles value labels, formats percentages, and 
structures output for easy interpretation.

{pstd}
Key features include:
{p 8 12}- Automatic detection of variable types (numeric/string){p_end}
{p 8 12}- Value label preservation and display{p_end}
{p 8 12}- Multiple display formats (frequencies, row/column/cell percentages){p_end}
{p 8 12}- Excel formatting (bold headers, number formats){p_end}
{p 8 12}- Progress tracking during export{p_end}
{p 8 12}- Comprehensive summary statistics{p_end}

{marker options}{...}
{title:Options}

{dlgtab:Main}

{phang}
{opt by(varlist)} specifies one or more variables for cross-tabulation. 
When specified, {cmd:exporttabs} creates two-way tables for each combination 
of variables in {it:varlist} with variables in {opt by()}. If omitted, 
one-way frequency tables are produced.

{phang}
{opt tabopt(string)} controls what values appear in table cells:
{pmore}
{opt row}: Display row percentages (percent of row total){p_end}
{pmore}
{opt col}: Display column percentages (percent of column total){p_end}
{pmore}
{opt cell}: Display cell percentages (percent of grand total){p_end}
{pmore}
If omitted, cells show frequencies (counts). Margins always show frequencies.

{phang}
{opt sheet(string)} specifies the Excel worksheet name. Default is "Tables".
Maximum 31 characters. Excel restrictions apply.

{dlgtab:Data handling}

{phang}
{opt missing} includes missing values ({cmd:.}, {cmd:.a}-{cmd:.z} for numeric; 
empty strings for string) in tabulations. By default, missing values are excluded.

{phang}
{opt nolabel} displays raw numeric values instead of value labels. Useful 
when value labels are too long or when exporting for statistical analysis.

{dlgtab:File handling}

{phang}
{opt force} overwrites existing Excel files without prompting for confirmation.
Use with caution as it permanently deletes existing files.

{phang}
{opt replace} replaces the specified sheet if it exists. By default, the 
program appends to existing sheets. Requires {cmd:force} to overwrite entire files.

{marker examples}{...}
{title:Examples}

{phang}{hline}
{pstd}Setup: Load example dataset{p_end}
{phang2}{cmd:. sysuse auto, clear}{p_end}

{phang}{hline}
{pstd}Example 1: Export all variables (one-way tables){p_end}
{phang2}{cmd:. exporttabs using "auto_tables.xlsx"}{p_end}

{phang}{hline}
{pstd}Example 2: Export specific variables{p_end}
{phang2}{cmd:. exporttabs make price mpg using "auto_summary.xlsx"}{p_end}

{phang}{hline}
{pstd}Example 3: Cross-tabulation with row percentages{p_end}
{phang2}{cmd:. exporttabs foreign using "cross_tabs.xlsx", by(rep78) tabopt(row)}{p_end}

{phang}{hline}
{pstd}Example 4: Multiple cross-tabulations{p_end}
{phang2}{cmd:. exporttabs price mpg using "analysis.xlsx", by(foreign rep78)}{p_end}

{phang}{hline}
{pstd}Example 5: Include missing values{p_end}
{phang2}{cmd:. exporttabs rep78 using "with_missing.xlsx", missing}{p_end}

{phang}{hline}
{pstd}Example 6: Custom sheet name{p_end}
{phang2}{cmd:. exporttabs using "report.xlsx", sheet("Auto Industry Analysis")}{p_end}

{phang}{hline}
{pstd}Example 7: With if condition{p_end}
{phang2}{cmd:. exporttabs price mpg using "domestic.xlsx" if foreign == 0}{p_end}

{phang}{hline}
{pstd}Example 8: Complex combination{p_end}
{phang2}{cmd:. exporttabs price mpg weight using "full_analysis.xlsx", by(foreign rep78) tabopt(col) if price > 5000}{p_end}

{marker advanced}{...}
{title:Advanced examples}

{phang}{hline}
{pstd}Batch processing multiple variables{p_end}
{phang2}{cmd:. local vars "price mpg weight length"}{p_end}
{phang2}{cmd:. exporttabs `vars' using "batch_output.xlsx", by(foreign)}{p_end}

{phang}{hline}
{pstd}Creating multiple reports with loops{p_end}
{phang2}{cmd:. foreach var in price mpg weight {c -(}}{p_end}
{phang2}{cmd:.     exporttabs `var' using "`var'_analysis.xlsx", by(foreign)}{p_end}
{phang2}{cmd:. {c )-}}{p_end}

{phang}{hline}
{pstd}Using temporary files{p_end}
{phang2}{cmd:. tempfile results}{p_end}
{phang2}{cmd:. exporttabs using "`results'", by(region)}{p_end}
{phang2}{cmd:. copy "`results'" "Final_Report.xlsx", replace}{p_end}

{marker remarks}{...}
{title:Remarks}

{pstd}
{bf:Variable and value labels:} The program automatically uses variable labels 
as table titles and value labels for category names. To view or modify labels:
{p 8 12}{cmd:. describe}{p_end}
{p 8 12}{cmd:. label list}{p_end}
{p 8 12}{cmd:. label variable varname "New label"}{p_end}
{p 8 12}{cmd:. label define lblname 1 "Yes" 2 "No"}{p_end}
{p 8 12}{cmd:. label values varname lblname}{p_end}

{pstd}
{bf:Excel formatting:} The program applies basic formatting (bold headers, 
number formats). Additional formatting (borders, colors, alignment) should 
be done manually in Excel. Percentages are formatted with two decimal places.

{pstd}
{bf:Limitations:}
{p 8 12}- Maximum Excel column width: 16,384 columns{p_end}
{p 8 12}- Maximum Excel row height: 1,048,576 rows{p_end}
{p 8 12}- Sheet names: Maximum 31 characters{p_end}
{p 8 12}- File paths: Use forward slashes or double backslashes{p_end}

{pstd}
{bf:Performance tips:}
{p 8 12}- Use {cmd:if/in} conditions to limit data size{p_end}
{p 8 12}- Process variables in batches for large datasets{p_end}
{p 8 12}- Close Excel before running to avoid file locking{p_end}
{p 8 12}- Use network-optimized paths for shared drives{p_end}

{marker saved_results}{...}
{title:Saved results}

{pstd}
{cmd:exporttabs} does not save results in Stata's memory. All output is 
written directly to the specified Excel file. The program displays:
{p 8 12}- Progress indicators for each table{p_end}
{p 8 12}- Total number of tables created{p_end}
{p 8 12}- File save confirmation{p_end}
{p 8 12}- Usage tips and reminders{p_end}

{marker error_messages}{...}
{title:Error messages}

{pstd}
Common errors and solutions:

{pmore}
{err:file filename.xlsx already exists}{p_end}
{pmore2}Solution: Use {cmd:force} option or specify different filename{p_end}

{pmore}
{err:invalid sheet name}{p_end}
{pmore2}Solution: Sheet names must be ≤31 characters, cannot contain : \ / ? * [ ] {p_end}

{pmore}
{err:insufficient memory}{p_end}
{pmore2}Solution: Use {cmd:if/in} to reduce dataset size or process in batches{p_end}

{pmore}
{err:Excel file locked}{p_end}
{pmore2}Solution: Close Excel or check file permissions{p_end}

{pmore}
{err:no observations}{p_end}
{pmore2}Solution: Check {cmd:if/in} conditions or use {cmd:missing} option{p_end}

{marker acknowledgments}{...}
{title:Acknowledgments}

{pstd}
The {cmd:exporttabs} program was developed to simplify the process of 
creating publication-ready tables from Stata. It builds upon Stata's 
built-in {cmd:tabulate} and {cmd:putexcel} functionality with additional 
formatting and automation features.

{marker also_see}{...}
{title:Also see}

{psee}
Manual: {manhelp tabulate R:tabulate}, {manhelp putexcel P:putexcel},
{manhelp export_excel P:export excel}

{psee}
User commands: {help tabout}, {help asdoc}, {help estout}

{marker citation}{...}
{title:Citation}

{pstd}
When using {cmd:exporttabs} in publications, please cite:

{phang}
Statistical Programming Team. 2026. exporttabs: Stata module to export 
frequency and cross-tabulation tables to Excel. Version 1.0.0.

{title:Author}

{pstd}
Statistical Programming Team{p_end}
{pstd}
Email: stats@example.com{p_end}
{pstd}
Website: {browse "https://www.example.com/stata-tools"}{p_end}

{title:Support}

{pstd}
For bug reports, feature requests, and technical support:{p_end}
{pstd}{browse "https://github.com/example/exporttabs/issues":GitHub Issues}{p_end}
{pstd}{browse "mailto:support@example.com":Email support}{p_end}

{title:Version}

{pstd}
Version 1.0.0 - 22 January 2026{p_end}
{pstd}
First public release{p_end}

{hline}
*/
