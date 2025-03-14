# main.py
import argparse
from utils.xml_converter import XMLConverter
from utils.excel_converter import ExcelConverter
from lxml import etree
from models.recipe import RecipeElement

def xml_to_excel(args):
    xml_conv = XMLConverter(xsd_path=args.xsd)
    recipe: RecipeElement = xml_conv.parse_xml(args.xml_input)
    print("Parsed XML into model:")
    print(recipe.model_dump_json(indent=2))
    
    excel_conv = ExcelConverter(args.excel_output)
    excel_conv.recipe_to_excel(recipe)

def excel_to_xml(args):
    xml_conv = XMLConverter(xsd_path=args.xsd)
    recipe: RecipeElement = xml_conv.parse_xml(args.xml_input)
    print("Parsed original XML into model:")
    print(recipe.model_dump_json(indent=2))
    
    excel_conv = ExcelConverter(args.excel_input_excel)
    # Update the recipe model with edited values from Excel
    recipe = excel_conv.excel_to_recipe(recipe)
    
    root = xml_conv.to_xml(recipe)
    tree = etree.ElementTree(root)
    tree.write(args.xml_output, pretty_print=True, xml_declaration=True, encoding="UTF-8")
    print(f"Modified XML saved to {args.xml_output}")

def parse_arguments():
    parser = argparse.ArgumentParser(
        description="CLI tool to convert Recipe XML to Excel and vice versa."
    )
    subparsers = parser.add_subparsers(dest="mode", required=True, help="Operation mode")

    # Subcommand: XML to Excel
    parser_xml2excel = subparsers.add_parser("xml2excel", help="Convert XML to Excel.")
    parser_xml2excel.add_argument("--xml-input", required=True, help="Path to the input XML file.")
    parser_xml2excel.add_argument("--excel-output", required=True, help="Path to the output Excel file.")
    parser_xml2excel.add_argument("--xsd", help="(Optional) Path to the XSD file for XML validation.")
    parser_xml2excel.set_defaults(func=xml_to_excel)

    # Subcommand: Excel to XML
    parser_excel2xml = subparsers.add_parser("excel2xml", help="Convert edited Excel back to XML.")
    parser_excel2xml.add_argument("--xml-input", required=True, help="Path to the original XML file (to retain the full model structure).")
    parser_excel2xml.add_argument("--excel-input-excel", required=True, help="Path to the edited Excel file.")
    parser_excel2xml.add_argument("--xml-output", required=True, help="Path to save the updated XML file.")
    parser_excel2xml.add_argument("--xsd", help="(Optional) Path to the XSD file for XML validation.")
    parser_excel2xml.set_defaults(func=excel_to_xml)

    return parser.parse_args()

def main():
    args = parse_arguments()
    args.func(args)

if __name__ == "__main__":
    main()
