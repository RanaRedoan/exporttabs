# optcounts

This repository contains a Stata program, `optcounts`, designed to analyze survey data by counting user-specified special value selections per enumerator.

## Files

- `optcounts.ado`: The Stata program file for counting special values (e.g., -99, 99, -999) across all variables per enumerator.
- `optcounts.sthlp`: The help file for Stata documentation.
- `example.do`: A detailed example script demonstrating usage with sample data.
- `readme.md`: This file, providing an overview of the repository.

## Installation

1. Clone or download this repository to your local machine.
2. Place the `optcounts.ado` file in your Stata ado path (run `sysdir` in Stata to find it, e.g., `~/ado/personal/`).
3. Install the help file by placing `optcounts.sthlp` in the same directory or your Stata help path.

## Usage

Run the following command in Stata after loading your dataset:

```stata
optcounts -99 99 -999 -98, enum(interviewr_name)
