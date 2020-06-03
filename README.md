# XLScolors

Colorize Excel spreadsheets according to keywords defined in a stylesheet, using the OpenPyxl library.


## Installing/Getting started

```shell
pip3 install -r requirements.txt
python xlscolors.py --infile my_excel_file.xlsx [--outfile colorized.xlsx]
```

Works with .xls and .xlsx files.


## Usage:
For available options, see :
```shell
python xlscolors.py --help
```

**ATTENTION:** if no ```--outfile``` is specified, the Excel workbook will be modified in place! Please backup your important files first.



## Stylesheet
xlscolors.py colors spreadsheets according to a stylesheet written in YAML :

- xlscolors will look for a .yaml file with the same name as the Excel file (e.g. file1.yaml for file1.xlsx)
- alternatively, a specific stylesheet can be specified at the command line with the --stylesheet argument
- if no file is found then it will use xlscolors.yaml
- if none of these files are found xlscolors will stop.



## Stylesheet Syntax

The stylesheet is plain and simple YAML, with the use of anchors and aliases to allow easy reuse of color names.

```yaml
  white: &white 'ffffff'
  black: &black '000000'
  red: &red 'ff0000' 
```

The stylesheet is divided in 3 sections:

### colors: Define base colors with their aliases.
```yaml
colors:
  #
  # Normal colors
  #
  white: &white 'ffffff'
  black: &black '000000'
```


### headers: define style for the header row (first row of each spreadsheet)
```yaml
headers:
  fg: *black
  bg: *yellow
  bold: true
```

### keywords: define foreground and background colors for cells matching a keyword
```yaml
keywords:
  'Yes':
    fg: *black
    bg: *green
  ++wrong++:
    fg: *black
    bg: *red
```

### Important :
- 'fg' and 'bg' objects must exist for each keyword
- some keywords like 'yes/no' and 'true/false' are reserved in YAML, surround them with quotes
- to match any cell _containing_ the keyword, surround it with '++'
- Matching is **NOT** case sensitive.

See the xlscolors.yaml file for more examples.


## Licensing

Author: David Paneels

This project is private and for internal use only. 
