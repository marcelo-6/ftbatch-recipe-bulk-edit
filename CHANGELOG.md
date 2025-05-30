## 0.2.1 (2025-05-29)

### Feat

- **logging**: added deferrals count to output
- **logging**: added better summary for what changes were made for excel2xml
- **xml_model**: support ParamExpression in FormulaValue
- auto-size columns to fit content
- **xml_model**: insert new <Parameter> nodes in correct position
- **parser**: added proper tracker of updated parameters
- **makefile**: add Make commands, making it easier to build / prepare to build. The pyinstaller still needs to run from windows

### Fix

- **xml_model**: fixed the way ParamExpression types are represented in excel
- **parser**: added save string strip for whern the excel sheet has boolean values (FALSE) for the scale cells for example
- **xml generation**: add extended debug logging
- **xml generation**: order of new parameters creation is now pre-determined

### Refactor

- reorganized code to make it easier to troubleshoot and read

## v0.1.4 (2025-05-19)

### Feat

- **makefile**: add Make commands, making it easier to build / prepare to...

## 0.2.0 (2025-05-21)

### Fix

- **xml generation**: add extended debug logging

### Refactor

- reorganized code to make it easier to troubleshoot and read

## 0.1.4 (2025-05-19)

### Fix

- **makefile**: exe will now generate metadata correctly

## 0.1.3 (2025-05-19)

### Fix

- removed vscode settings from git

## 0.1.2 (2025-05-19)

### Feat

- **makefile**: add Make commands, making it easier to build / prepare to build. The pyinstaller still needs to run from windows

## 0.1.1 (2025-05-19)

### Fix

- **xml generation**: order of new parameters creation is now pre-determined

## 0.1.0 (2025-03-16)

### BREAKING CHANGE

- new dependencies

### Feat

- added first pass for xml2excel and excel2xml code

### Fix

- edge case for no argument
- fixed multiple file parse based on parent file
- added better logging
- better parse for a single file with complete conversion of oxml

## 0.0.2 (2025-02-08)

### Fix

- **Global**: Migrated pre-commit to latest
- **Global**: Added details to pyproject.toml along with configuration for python semantic versioning release
