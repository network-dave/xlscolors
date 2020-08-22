# XLScolors

Colorize Excel spreadsheets according to keywords defined in a stylesheet, using the OpenPyxl library.


## Requirements / Getting started

```shell
$ pip3 install -r requirements.txt
$ python xlscolors.py --infile my_excel_file.xlsx [--outfile colorized.xlsx]
```

Works with .xls and .xlsx files.


## Usage:

For available options, see :
```shell
python xlscolors.py --help
```

**ATTENTION:** if no ```--outfile``` is specified, the Excel workbook will be modified in place! Please backup your important files just in case.



## Stylesheet
xlscolors.py colors spreadsheets according to a stylesheet written in YAML :

- xlscolors will look for a .yaml file with the same name as the Excel file (e.g. file1.yaml for file1.xlsx)
- alternatively, a specific stylesheet can be specified at the command line with the --stylesheet argument
- if no file is found then it will look for xlscolors.yaml in the currect working directory, then in xlscolors.py's directory
- if none of these files are found then xlscolors will exit.



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


## License

Copyright (C) 2020 David Paneels

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

https://www.gnu.org/licenses/gpl-3.0.html
