# Excel Tricks and VBA Codes

A collection of Excel VBA macros that automate repetitive spreadsheet tasks.

## Macros

| Macro | File | What it does |
|---|---|---|
| Copy and rename sheets | `copy-and-rename-sheets.bas` | Creates one copy of a template sheet for every name in a list, and names each copy automatically |

## Copy and rename sheets

Useful when you need the same sheet layout many times, for example one sheet per month, branch or employee.

**Setup**

1. Put the layout you want to copy in a sheet named `Sheet1`.
2. List the new sheet names in column A of a sheet named `Sheet2`, one name per row.

**Run it**

1. Open the Developer tab and click Visual Basic.
2. Choose Insert, then Module.
3. Paste the code from `copy-and-rename-sheets.bas`.
4. Run `CopySheet`. One sheet is created for each name in the list.

Sheet names must be unique and must follow Excel's naming rules (31 characters at most, and none of `\ / ? * [ ] :`).

## Tools

Microsoft Excel · VBA

## Author

**Omer Naeem Khan** — Data Analyst
[GitHub](https://github.com/OmerNaeemKhan) · [LinkedIn](https://www.linkedin.com/in/omer-khan-03833749)
