# XML Recipe Defferal Converter CLI Tool

This project provides a command-line tool to convert complex XML files (with a known XSD schema) into Excel spreadsheets and back into XML. It uses a class-based design with:

- **Pydantic** for data validation,
- **lxml** for XML parsing/generation, and
- **openpyxl** for Excel operations.

Engineers can use this tool in two modes:
- **xml2excel**: Convert an XML file to a multi-sheet Excel file for bulk editing.
- **excel2xml**: Convert the modified Excel file back into an XML file.

## Project Structure

```
converter/
├── main.py
├── models/
│   ├── __init__.py
│   └── recipe.py
├── utils/
│   ├── __init__.py
│   ├── xml_converter.py
│   └── excel_converter.py
├── pyproject.toml
└── README.md
```

- **models/recipe.py**: Contains the Pydantic models representing the XML schema.
- **utils/xml_converter.py**: Provides functionality to parse XML into models and convert models back to XML.
- **utils/excel_converter.py**: Provides functionality to export the full XML data into a multi-sheet Excel file and import changes back into the models.
- **main.py**: The CLI entry point that supports subcommands for converting in both directions.
- **pyproject.toml**: The project configuration file used with UV for dependency management.

## Getting Started with UV

This project uses [uv](https://docs.astral.sh/uv) as the package manager. Follow the steps below to set up your virtual environment, install dependencies, and run the CLI tool.

### 1. Create a Virtual Environment

Make sure you have UV installed. Create a virtual environment with the desired Python version (e.g., Python 3.10):

```bash
uv venv --python 3.10
```

Activate your virtual environment:

- **On Unix or macOS:**

  ```bash
  source .venv/bin/activate
  ```

- **On Windows:**

  ```bash
  .venv\Scripts\activate
  ```

### 2. Install Dependencies

The project dependencies are defined in the `pyproject.toml` file. To install them, run:

```bash
uv pip install -r pyproject.toml
```

This command will install all required packages (such as `pydantic`, `lxml`, and `openpyxl`).

### 3. Running the CLI Tool

The CLI tool is invoked via the `main.py` file and supports two subcommands:

#### Convert XML to Excel

This command parses an input XML file and exports its contents to an Excel file with multiple sheets.

```bash
uv run main.py xml2excel --xml-input <path_to_input_xml> --excel-output <path_to_output_excel> --xsd <path_to_xsd>
```

**Example:**

```bash
uv run main.py xml2excel --xml-input sample_recipe.xml --excel-output full_recipe.xlsx --xsd MasterRecipe.xsd
```

#### Convert Excel to XML

After editing the Excel file, this command reads the updated Excel file and generates a new XML file.

```bash
uv run main.py excel2xml --xml-input <path_to_original_xml> --excel-input-excel <path_to_edited_excel> --xml-output <path_to_output_xml> --xsd <path_to_xsd>
```

**Example:**

```bash
uv run main.py excel2xml --xml-input sample_recipe.xml --excel-input-excel full_recipe.xlsx --xml-output updated_recipe.xml --xsd MasterRecipe.xsd
```

## How It Works

1. **XML Parsing and Model Mapping**  
   The `XMLConverter` class in `utils/xml_converter.py` reads the XML file and maps its contents to the Pydantic models defined in `models/recipe.py`. XML validation is performed against an optional XSD file if provided.

2. **Excel Export and Import**  
   The `ExcelConverter` class in `utils/excel_converter.py` exports the entire XML content into a multi-sheet Excel file, with separate sheets for Header, Parameters, Steps, etc. Engineers can edit the Excel file, and the tool can import these modifications back into the models to regenerate the XML.

3. **Command-Line Interface**  
   The `main.py` file uses Python’s built-in `argparse` module to provide a CLI interface with subcommands for converting XML to Excel and vice versa.