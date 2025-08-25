//using Dataset
u "Dataset.dta", clear

// keeping relevent variables
keep district incomegroup age group sex education income type health_issue 
// or you can use varlist directly within exporttabs [varlist] using "tables.xlsx"
// export single tables

exporttabs using "01 output_single.xlsx"
exporttabs using "02 output_corss_freq.xlsx", by(district)
exporttabs using "03 output_col.xlsx", by(district) tabopt("col")
exporttabs using "04 output_col_nofreq.xlsx", by(district) tabopt("col nofreq")
exporttabs using "05 output_row.xlsx", by(district) tabopt("row")
exporttabs using "06 output_row_nofreq.xlsx", by(district) tabopt("row nofreq")
exporttabs using "07 output_cell.xlsx", by(district) tabopt(cell)
