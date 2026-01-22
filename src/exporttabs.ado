*! version 1.0.3 31aug2025
*! Author: Md. Redoan Hossain Bhuiyan
*! Program: exporttabs
*! Purpose: Export one-way and two-way tabulations to Excel in batch (fixed value label issue)

program define exporttabs
    version 15.0
    syntax [varlist(default=none)] using/, [BY(varlist) TABOPT(string)]

    * If no varlist specified, use all variables
    if "`varlist'" == "" {
        unab varlist : _all
    }

    * Parse crosstab display options
    local optlower = lower("`tabopt'")
    local want_row  = strpos(" `optlower' ", " row ")  > 0
    local want_col  = strpos(" `optlower' ", " col ")  > 0
    local want_cell = strpos(" `optlower' ", " cell ") > 0
    local want_pct  = `want_row' | `want_col' | `want_cell'

    * Initialize Excel file - quietly suppress save messages
    quietly putexcel set "`using'", sheet("Tables") replace

    local row = 1
    local n_tables = 0

    * Add report header - quietly suppress messages
    quietly {
        putexcel A`row' = ("TABULATION REPORT"), bold
        local ++row
        putexcel A`row' = ("Generated: `c(current_date)' `c(current_time)'")
        local row = `row' + 2
    }

    * Display progress header
    di as txt "{hline 70}"
    di as txt "{bf:TABULATION EXPORT IN PROGRESS...}"
    di as txt "{hline 70}"

    * Loop through each variable
    foreach v of varlist `varlist' {
        
        * Get variable label
        local vlabel : variable label `v'
        if "`vlabel'" == "" local vlabel "`v'"

        * Check if variable is string or numeric
        capture confirm numeric variable `v'
        local is_numeric = !_rc

        * ONE-WAY TABULATION (no by variable)
        if "`by'" == "" {
            
            * Show progress message
            di as txt "  → Processing: " as result "`v'" as txt " ..." _continue
            
            * Check if variable has valid observations
            if `is_numeric' {
                quietly count if !missing(`v')
            }
            else {
                quietly count if `v' != ""
            }
            
            if r(N) == 0 {
                di as error " SKIPPED (no data)"
                continue
            }

            * Write variable header
            quietly putexcel A`row' = ("`v' (`vlabel')"), bold
            local ++row

            * Get levels
            if `is_numeric' {
                quietly levelsof `v' if !missing(`v'), local(levels)
                quietly count if !missing(`v')
            }
            else {
                quietly levelsof `v' if `v' != "", local(levels)
                quietly count if `v' != ""
            }
            local total_n = r(N)

            * Write column headers
            quietly putexcel A`row' = ("Value") B`row' = ("Frequency") C`row' = ("Percent"), bold
            local ++row

            * Get value label if exists
            local vallab : value label `v'

            * Write each level - quietly suppress messages
            quietly {
                foreach lev of local levels {
                    * Get label for this level
                    if "`vallab'" != "" & `is_numeric' {
                        capture local lbl : label (`vallab') `lev'
                        if _rc local lbl "`lev'"
                    }
                    else {
                        local lbl "`lev'"
                    }

                    * Count frequency
                    if `is_numeric' {
                        count if `v' == `lev'
                    }
                    else {
                        count if `v' == "`lev'"
                    }
                    local freq = r(N)
                    local pct = round(100 * `freq' / `total_n', 0.01)

                    * Write to Excel
                    putexcel A`row' = ("`lbl'")
                    putexcel B`row' = (`freq')
                    putexcel C`row' = (`pct'), nformat(number_d2)
                    local ++row
                }

                * Write total row
                putexcel A`row' = ("Total"), bold
                putexcel B`row' = (`total_n'), bold
                putexcel C`row' = (100), bold nformat(number_d2)
                local row = `row' + 2
                local ++n_tables
            }
            
            * Show completion
            di as result " ✓ COMPLETED"
        }
        
        * TWO-WAY TABULATION (with by variable)
        else {
            foreach a of varlist `by' {
                
                * Skip if same variable
                if "`v'" == "`a'" continue

                * Show progress message
                di as txt "  → Processing: " as result "`v' × `a'" as txt " ..." _continue

                * Check if by-variable is string or numeric
                capture confirm numeric variable `a'
                local a_is_numeric = !_rc

                * Get by-variable label
                local alabel : variable label `a'
                if "`alabel'" == "" local alabel "`a'"

                * Check if enough valid observations
                if `is_numeric' & `a_is_numeric' {
                    quietly count if !missing(`v') & !missing(`a')
                }
                else if `is_numeric' & !`a_is_numeric' {
                    quietly count if !missing(`v') & `a' != ""
                }
                else if !`is_numeric' & `a_is_numeric' {
                    quietly count if `v' != "" & !missing(`a')
                }
                else {
                    quietly count if `v' != "" & `a' != ""
                }
                
                if r(N) == 0 {
                    di as error " SKIPPED (no data)"
                    continue
                }

                * Write table header - quietly suppress messages
                quietly {
                    putexcel A`row' = ("`v' (`vlabel') × `a' (`alabel')"), bold
                    local ++row

                    * Get levels for both variables
                    if `is_numeric' & `a_is_numeric' {
                        levelsof `v' if !missing(`v') & !missing(`a'), local(rlevels)
                        levelsof `a' if !missing(`v') & !missing(`a'), local(clevels)
                    }
                    else if `is_numeric' & !`a_is_numeric' {
                        levelsof `v' if !missing(`v') & `a' != "", local(rlevels)
                        levelsof `a' if !missing(`v') & `a' != "", local(clevels)
                    }
                    else if !`is_numeric' & `a_is_numeric' {
                        levelsof `v' if `v' != "" & !missing(`a'), local(rlevels)
                        levelsof `a' if `v' != "" & !missing(`a'), local(clevels)
                    }
                    else {
                        levelsof `v' if `v' != "" & `a' != "", local(rlevels)
                        levelsof `a' if `v' != "" & `a' != "", local(clevels)
                    }
                }

                * Count rows and columns
                local rN : word count `rlevels'
                local cN : word count `clevels'

                * Get value labels
                local vallab_r : value label `v'
                local vallab_c : value label `a'

                * Calculate all frequencies and totals
                tempname freq_mat rtot ctot
                matrix `freq_mat' = J(`rN', `cN', 0)
                matrix `rtot' = J(`rN', 1, 0)
                matrix `ctot' = J(1, `cN', 0)
                local grandtotal = 0

                local i = 1
                foreach rl of local rlevels {
                    local j = 1
                    foreach cl of local clevels {
                        * Count based on variable types
                        if `is_numeric' & `a_is_numeric' {
                            quietly count if `v' == `rl' & `a' == `cl'
                        }
                        else if `is_numeric' & !`a_is_numeric' {
                            quietly count if `v' == `rl' & `a' == "`cl'"
                        }
                        else if !`is_numeric' & `a_is_numeric' {
                            quietly count if `v' == "`rl'" & `a' == `cl'
                        }
                        else {
                            quietly count if `v' == "`rl'" & `a' == "`cl'"
                        }
                        
                        local freq = r(N)
                        matrix `freq_mat'[`i',`j'] = `freq'
                        matrix `rtot'[`i',1] = `rtot'[`i',1] + `freq'
                        matrix `ctot'[1,`j'] = `ctot'[1,`j'] + `freq'
                        local grandtotal = `grandtotal' + `freq'
                        local ++j
                    }
                    local ++i
                }

                * Write column headers - quietly suppress messages
                quietly {
                    putexcel A`row' = ("Value"), bold
                    local j = 1
                    foreach cl of local clevels {
                        * Get label for column
                        if "`vallab_c'" != "" & `a_is_numeric' {
                            capture local clbl : label (`vallab_c') `cl'
                            if _rc local clbl "`cl'"
                        }
                        else {
                            local clbl "`cl'"
                        }
                        
                        * Calculate Excel column (handle beyond Z)
                        local col_num = `j' + 1
                        if `col_num' <= 26 {
                            local col_let : word `col_num' of `c(ALPHA)'
                        }
                        else {
                            local c1 = int((`col_num'-1)/26)
                            local c2 = mod(`col_num'-1, 26) + 1
                            local col_let1 : word `c1' of `c(ALPHA)'
                            local col_let2 : word `c2' of `c(ALPHA)'
                            local col_let "`col_let1'`col_let2'"
                        }
                        
                        putexcel `col_let'`row' = ("`clbl'"), bold
                        local ++j
                    }
                    
                    * Add Total column header
                    local col_num = `cN' + 2
                    if `col_num' <= 26 {
                        local col_let : word `col_num' of `c(ALPHA)'
                    }
                    else {
                        local c1 = int((`col_num'-1)/26)
                        local c2 = mod(`col_num'-1, 26) + 1
                        local col_let1 : word `c1' of `c(ALPHA)'
                        local col_let2 : word `c2' of `c(ALPHA)'
                        local col_let "`col_let1'`col_let2'"
                    }
                    putexcel `col_let'`row' = ("Total (N)"), bold
                    local ++row

                    * Write data rows
                    local i = 1
                    foreach rl of local rlevels {
                        * Get row label
                        if "`vallab_r'" != "" & `is_numeric' {
                            capture local rlbl : label (`vallab_r') `rl'
                            if _rc local rlbl "`rl'"
                        }
                        else {
                            local rlbl "`rl'"
                        }
                        
                        putexcel A`row' = ("`rlbl'")

                        * Write each cell
                        local j = 1
                        foreach cl of local clevels {
                            local freq = `freq_mat'[`i',`j']
                            local cell_value = .
                            
                            * Calculate cell value based on options
                            if `want_row' {
                                local denom = `rtot'[`i',1]
                                if `denom' > 0 {
                                    local cell_value = round(100 * `freq' / `denom', 0.01)
                                }
                            }
                            else if `want_col' {
                                local denom = `ctot'[1,`j']
                                if `denom' > 0 {
                                    local cell_value = round(100 * `freq' / `denom', 0.01)
                                }
                            }
                            else if `want_cell' {
                                if `grandtotal' > 0 {
                                    local cell_value = round(100 * `freq' / `grandtotal', 0.01)
                                }
                            }
                            else {
                                local cell_value = `freq'
                            }

                            * Calculate Excel column
                            local col_num = `j' + 1
                            if `col_num' <= 26 {
                                local col_let : word `col_num' of `c(ALPHA)'
                            }
                            else {
                                local c1 = int((`col_num'-1)/26)
                                local c2 = mod(`col_num'-1, 26) + 1
                                local col_let1 : word `c1' of `c(ALPHA)'
                                local col_let2 : word `c2' of `c(ALPHA)'
                                local col_let "`col_let1'`col_let2'"
                            }
                            
                            if `want_pct' {
                                putexcel `col_let'`row' = (`cell_value'), nformat(number_d2)
                            }
                            else {
                                putexcel `col_let'`row' = (`cell_value')
                            }
                            local ++j
                        }

                        * Write row total
                        local col_num = `cN' + 2
                        if `col_num' <= 26 {
                            local col_let : word `col_num' of `c(ALPHA)'
                        }
                        else {
                            local c1 = int((`col_num'-1)/26)
                            local c2 = mod(`col_num'-1, 26) + 1
                            local col_let1 : word `c1' of `c(ALPHA)'
                            local col_let2 : word `c2' of `c(ALPHA)'
                            local col_let "`col_let1'`col_let2'"
                        }
                        putexcel `col_let'`row' = (`rtot'[`i',1])
                        local ++row
                        local ++i
                    }

                    * Write column totals row
                    putexcel A`row' = ("Total (N)"), bold
                    forvalues j = 1/`cN' {
                        local col_num = `j' + 1
                        if `col_num' <= 26 {
                            local col_let : word `col_num' of `c(ALPHA)'
                        }
                        else {
                            local c1 = int((`col_num'-1)/26)
                            local c2 = mod(`col_num'-1, 26) + 1
                            local col_let1 : word `c1' of `c(ALPHA)'
                            local col_let2 : word `c2' of `c(ALPHA)'
                            local col_let "`col_let1'`col_let2'"
                        }
                        putexcel `col_let'`row' = (`ctot'[1,`j']), bold
                    }
                    
                    * Write grand total
                    local col_num = `cN' + 2
                    if `col_num' <= 26 {
                        local col_let : word `col_num' of `c(ALPHA)'
                    }
                    else {
                        local c1 = int((`col_num'-1)/26)
                        local c2 = mod(`col_num'-1, 26) + 1
                        local col_let1 : word `c1' of `c(ALPHA)'
                        local col_let2 : word `c2' of `c(ALPHA)'
                        local col_let "`col_let1'`col_let2'"
                    }
                    putexcel `col_let'`row' = (`grandtotal'), bold
                    local ++row

                    * Add note if percentages shown
                    if `want_pct' {
                        local how = cond(`want_row', "row", cond(`want_col', "column", "cell"))
                        putexcel A`row' = ("Note: cells show `how' percentages; margins show counts (N).")
                        local ++row
                    }

                    local row = `row' + 1
                    local ++n_tables
                }
                
                * Show completion
                di as result " ✓ COMPLETED"
            }
        }
    }

    * Final success message
    di as txt "{hline 70}"
    di as txt  "                 " as result "✔ EXPORT COMPLETED SUCCESSFULLY ✔"
    di as txt  "{hline 65}"
    di as txt  "   Number of tables created : " as res `n_tables'
    di as txt  "   File saved as            : " as res "`using'"
    di as txt  "{hline 65}"
    di as txt  "   TIPS:"
    di as txt  "     • Use {bf:by()} for cross-tabs"
    di as txt  "     • Use {bf:tabopt(row|col|cell)} to control cell display"
    di as txt  "     • Manually format Excel tables (borders, shading, fonts)"
    di as txt  "     • Percentages are rounded to 2 decimals"
    di as txt  "     • Always check totals (N) when interpreting percentages"
    di as txt  "{hline 65}"
    di as txt  "        Thank you for using " as result "exporttabs" as txt "!"
    di as txt  "{hline 65}"
end
