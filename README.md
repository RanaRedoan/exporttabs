# exporttabs: Export Tabulations to Excel from Stata

`exporttabs` is a Stata program that exports single and cross tabulations to **Excel** in a clean, ready-to-use format.  
It supports flexible options for **row/column/cell percentages** and can batch-process all variables in your dataset.

---

## 🔧 Installation

Clone or download this repository and place `exporttabs.ado` and `exporttabs.sthlp` in your Stata `ado` path.

```stata
* Example: install directly from GitHub (if you use github integration)
net install exporttabs, from("https://raw.githubusercontent.com/RanaRedoan/exporttabs/main") replace
```

---

## 📖 Syntax

```stata
exporttabs [varlist] using filename.xlsx , [ by(varlist) tabopt(string) ]
```

---

## 📌 Options

```text
by(varlist)
    Create crosstabs with one or more variables.
    Example: by(district)

tabopt(string)
    Pass tabulation options such as:
        col     → Column percentages
        row     → Row percentages
        cell    → Cell percentages
        nofreq  → Suppress frequencies
```

---

## 📊 Examples

Suppose you have survey data with 250 respondents across **5 districts**:  
Dhaka, Cumilla, Chandpur, Gazipur, Cox's Bazar.  
Variable `age_group` has the age categories.

```stata
* 1. Single variable tabulation
exporttabs using "01 out_single.xlsx"

* 2. Crosstab with frequencies
exporttabs using "02 out_cross_freq.xlsx", by(district)

* 3. Column percentages
exporttabs using "03 out_col.xlsx", by(district) tabopt("col")

* 4. Column percentages without frequencies
exporttabs using "04 out_col_nofreq.xlsx", by(district) tabopt("col nofreq")

* 5. Row percentages
exporttabs using "05 out_row.xlsx", by(district) tabopt("row")

* 6. Row percentages without frequencies
exporttabs using "06 out_row_nofreq.xlsx", by(district) tabopt("row nofreq")

* 7. Cell percentages
exporttabs using "07 out_cell.xlsx", by(district) tabopt("cell")
```

---

## ✅ Output

```text
- All results are exported into the specified Excel file.  
- Each table includes labels, frequencies/percentages, and totals.  
- The output is raw but clean — users can apply Excel formatting as desired.  
- (Optional) You can also maintain a pre-formatted Excel template in your repo 
  and adapt the results into it.
```

---

## 📌 Notes & Tips

```text
- Use by() for cross tabulations.  
- Use tabopt() to control whether frequencies, row %, col %, or cell % appear.  
- All tables are plain-format to maximize compatibility.  
- For better visuals, add borders, adjust alignment, and apply Excel formatting.
```

---

## 🤝 Contribution

```text
Pull requests and suggestions are welcome!  
If you find issues or have feature requests, please open an Issue in the repository.
```

---

## 📜 License

```text
This project is licensed under the MIT License.
```

