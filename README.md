<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>exporttabs: Export Tabulations to Excel from Stata</title>
<style>
    body {
        font-family: Arial, sans-serif;
        line-height: 1.6;
        background-color: #f9f9f9;
        color: #333;
        padding: 20px;
    }
    h1, h2, h3 {
        color: #2c3e50;
    }
    h1 {
        font-size: 2em;
    }
    h2 {
        border-bottom: 2px solid #2c3e50;
        padding-bottom: 5px;
    }
    pre {
        background-color: #eee;
        padding: 10px;
        overflow-x: auto;
        border-radius: 5px;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin: 15px 0;
    }
    table, th, td {
        border: 1px solid #bbb;
    }
    th, td {
        padding: 10px;
        text-align: left;
    }
    code {
        background-color: #eee;
        padding: 2px 5px;
        border-radius: 3px;
    }
    .emoji {
        font-size: 1.2em;
    }
</style>
</head>
<body>

<h1>🚀 exporttabs: Export Tabulations to Excel from Stata</h1>
<p><code>exporttabs</code> is a <strong>Stata program</strong> designed to export single and cross tabulations directly to <strong>Excel</strong> in a clean, ready-to-use format.</p>
<p>It supports flexible options for <strong>row, column, or cell percentages</strong> and can batch-process all variables in your dataset effortlessly.</p>

<h2>🔧 Installation</h2>
<p>Clone or download this repository and place the files <code>exporttabs.ado</code> and <code>exporttabs.sthlp</code> in your <strong>Stata <code>ado</code> path</strong>.</p>

<pre><code>* Install directly from GitHub
net install exporttabs, from("https://raw.githubusercontent.com/RanaRedoan/exporttabs/main") replace
</code></pre>

<h2>📖 Syntax</h2>
<pre><code>exporttabs [varlist] using filename.xlsx , [ by(varlist) tabopt(string) ]</code></pre>

<h2>📌 Options</h2>
<table>
    <thead>
        <tr>
            <th>Option</th>
            <th>Description</th>
            <th>Example</th>
        </tr>
    </thead>
    <tbody>
        <tr>
            <td><code>by(varlist)</code></td>
            <td>Create cross-tabulations with one or more variables</td>
            <td><code>by(district)</code></td>
        </tr>
        <tr>
            <td><code>tabopt(string)</code></td>
            <td>Control tabulation output</td>
            <td>
                <code>"col"</code> → Column %<br>
                <code>"row"</code> → Row %<br>
                <code>"cell"</code> → Cell %<br>
                <code>"nofreq"</code> → Suppress frequencies
            </td>
        </tr>
    </tbody>
</table>

<h2>📊 Examples</h2>
<p>Suppose your survey dataset contains <strong>250 respondents across 5 districts</strong>: Dhaka, Cumilla, Chandpur, Gazipur, Cox's Bazar. Variable <code>age_group</code> represents age categories.</p>

<pre><code>* 1️⃣ Single variable tabulation
exporttabs using "01_out_single.xlsx"

* 2️⃣ Crosstab with frequencies
exporttabs using "02_out_cross_freq.xlsx", by(district)

* 3️⃣ Column percentages
exporttabs using "03_out_col.xlsx", by(district) tabopt("col")

* 4️⃣ Column percentages without frequencies
exporttabs using "04_out_col_nofreq.xlsx", by(district) tabopt("col nofreq")

* 5️⃣ Row percentages
exporttabs using "05_out_row.xlsx", by(district) tabopt("row")

* 6️⃣ Row percentages without frequencies
exporttabs using "06_out_row_nofreq.xlsx", by(district) tabopt("row nofreq")

* 7️⃣ Cell percentages
exporttabs using "07_out_cell.xlsx", by(district) tabopt("cell")
</code></pre>

<h2>✅ Output</h2>
<ul>
    <li>All tables are exported to the specified <strong>Excel file</strong>.</li>
    <li>Each table includes <strong>labels, frequencies/percentages, and totals</strong>.</li>
    <li>Output is <strong>clean and raw</strong>, allowing users to apply Excel formatting as needed.</li>
    <li>Optional: Maintain a pre-formatted Excel template and adapt results directly.</li>
</ul>

<h2>💡 Notes & Tips</h2>
<ul>
    <li>Use <code>by()</code> to create <strong>cross-tabulations</strong>.</li>
    <li>Use <code>tabopt()</code> to control display of <strong>frequencies, row %, column %, or cell %</strong>.</li>
    <li>Tables are <strong>plain-format</strong> for maximum compatibility.</li>
    <li>For enhanced visuals, apply <strong>borders, alignment adjustments, and Excel formatting</strong>.</li>
</ul>

<h2>🤝 Contributing</h2>
<p>Pull requests and suggestions are welcome! If you find issues or have feature requests, please open an <strong>Issue</strong> in the repository.</p>

<h2>📜 License</h2>
<p>This project is licensed under the <strong>MIT License</strong>.</p>

</body>
</html>
