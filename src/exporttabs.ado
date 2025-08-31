*! version 1.0.3 31aug2025
*! Author: Md. Redoan Hossain Bhuiyan
*! Program: exporttabs
*! Purpose: Export one-way and two-way tabulations to Excel in batch (fixed value label issue)

program define exporttabs
    version 15.0
    syntax [varlist(default=none)] using/, [BY(varlist) TABOPT(string)]

    if "`varlist'" == "" {
        unab varlist : _all
    }

    // ----------------------------
    // Ensure each variable has a value label with the same name
    ds
    local allvars `r(varlist)'
    foreach var of local allvars {
        local vlabel : value label `var'
        if "`vlabel'" != "" {
            capture label copy `vlabel' `var'
            label values `var' `var'
        }
    }
    // ----------------------------

    // Parse crosstab display options
    local optlower = lower("`tabopt'")
    local want_row  = strpos(" `optlower' ", " row ")  > 0
    local want_col  = strpos(" `optlower' ", " col ")  > 0
    local want_cell = strpos(" `optlower' ", " cell ") > 0
    local want_pct  = `want_row' | `want_col' | `want_cell'

    putexcel set "`using'", sheet("Tables") replace

    local row = 1
    local n_tables = 0

    foreach v of varlist `varlist' {
        local vlabel : variable label `v'
        if "`vlabel'" == "" local vlabel "`v'"

        if "`by'" == "" {
            putexcel A`row' = "`v' (`vlabel')", bold
            local ++row

            quietly levelsof `v', local(levels)
            quietly tab `v', matcell(freq)

            putexcel A`row' = "Value" B`row' = "Frequency" C`row' = "Percent", bold
            local ++row

            local vallab : value label `v'
            local i = 1
            scalar total = 0
            foreach l of local levels {
                local lbl : label (`vallab') `l'
                putexcel A`row' = "`lbl'"

                scalar f = freq[`i',1]
                scalar p = round(100*f/r(N), .01)
                putexcel B`row' = f C`row' = p
                scalar total = total + f
                local ++row
                local ++i
            }

            putexcel A`row' = "Total", bold
            putexcel B`row' = total, bold
            putexcel C`row' = 100, bold
            local row = `row' + 2
            local n_tables = `n_tables' + 1
        }
        else {
            foreach a of varlist `by' {
                if "`v'" != "`a'" {
                    local alabel : variable label `a'
                    if "`alabel'" == "" local alabel "`a'"

                    putexcel A`row' = "`v' (`vlabel') × `a' (`alabel')", bold
                    local ++row

                    quietly tab `v' `a', matcell(freq) `tabopt'

                    local rN = rowsof(freq)
                    local cN = colsof(freq)
                    quietly levelsof `v' if !missing(`v') & !missing(`a'), local(rlevels)
                    quietly levelsof `a' if !missing(`v') & !missing(`a'), local(clevels)

                    local vallab_r : value label `v'
                    local vallab_c : value label `a'

                    tempname rtot ctot
                    matrix `rtot' = J(`rN',1,0)
                    matrix `ctot' = J(1,`cN',0)
                    scalar grandtotal = 0
                    forvalues i = 1/`rN' {
                        forvalues j = 1/`cN' {
                            scalar f = freq[`i',`j']
                            matrix `rtot'[`i',1] = `rtot'[`i',1] + f
                            matrix `ctot'[1,`j'] = `ctot'[1,`j'] + f
                            scalar grandtotal = grandtotal + f
                        }
                    }

                    putexcel A`row' = "Value", bold
                    local j = 1
                    foreach cl of local clevels {
                        local clbl : label (`vallab_c') `cl'
                        putexcel `=char(65+`j')'`row' = "`clbl'", bold
                        local ++j
                    }
                    putexcel `=char(65+`cN'+1)'`row' = "Total (N)", bold
                    local ++row

                    local i = 1
                    foreach rl of local rlevels {
                        local rlbl : label (`vallab_r') `rl'
                        putexcel A`row' = "`rlbl'"

                        local j = 1
                        foreach cl of local clevels {
                            scalar f = freq[`i',`j']
                            scalar cell = .
                            if `want_row' {
                                scalar denom = `rtot'[`i',1]
                                if denom>0 scalar cell = round(100*f/denom,.01)
                            }
                            else if `want_col' {
                                scalar denom = `ctot'[1,`j']
                                if denom>0 scalar cell = round(100*f/denom,.01)
                            }
                            else if `want_cell' {
                                if grandtotal>0 scalar cell = round(100*f/grandtotal,.01)
                            }
                            else {
                                scalar cell = f
                            }
                            putexcel `=char(65+`j')'`row' = cell
                            local ++j
                        }

                        putexcel `=char(65+`cN'+1)'`row' = `rtot'[`i',1]
                        local ++row
                        local ++i
                    }

                    putexcel A`row' = "Total (N)", bold
                    forvalues j = 1/`cN' {
                        putexcel `=char(65+`j')'`row' = `ctot'[1,`j'], bold
                    }
                    putexcel `=char(65+`cN'+1)'`row' = grandtotal, bold
                    local ++row

                    if `want_pct' {
                        local how = cond(`want_row',"row", cond(`want_col',"column","cell"))
                        putexcel A`row' = "Note: cells show `how' percentages; margins show counts (N)."
                        local ++row
                    }

                    local row = `row' + 1
                    local n_tables = `n_tables' + 1
                }
            }
        }
    }

    // Final message
    di as txt  "{hline 65}"
    di as txt  "                 " as result "✔ EXPORT COMPLETED SUCCESSFULLY ✔"
    di as txt  "{hline 65}"
    di as txt  "   Number of tables created : " as res `n_tables'
    di as txt  "   File saved as            : " as res "`using'"
    di as txt  "{hline 65}"
    di as txt  "   TIPs:"
    di as txt  "     • Use {bf:by()} for cross-tabs"
    di as txt  "     • Use {bf:tabopt(row|col|cell [nofreq])} to control cell display"
    di as txt  "     • Manually format Excel tables (borders, shading, fonts)"
    di as txt  "     • Percentages are rounded to 2 decimals – adjust in Excel if needed"
    di as txt  "     • Always check totals (N) when interpreting percentages"
    di as txt  "     • For large surveys: filter or use {bf:if/in} to limit tables"
    di as txt  "{hline 65}"
    di as txt  "        Thank you for using " as result "exporttabs" as txt "!"
    di as txt  "{hline 65}"
end
