"""
XML model: RecipeTree, ParameterNode, FormulaValueNode.
"""

from lxml import etree
from lxml.etree import QName
from core.base import NAMESPACE, NSMAP, EXCEL_COLUMNS


class NodeBase:
    """
    Base for ParameterNode and FormulaValueNode.
    Holds the original XML element, fullpath, and metadata.
    """

    def __init__(self, element: etree.Element, fullpath: str, source_file: str):
        self.element = element
        self.fullpath = fullpath
        self.source_file = source_file
        # capture all sub-elements as original text ("" if empty)
        self.original_subs = {
            QName(child.tag).localname: (child.text or "") for child in element
        }

    def to_excel_row(self) -> dict:
        """
        Build a dict mapping Excel column names -> values for this node.
        """
        # TODO Implement me bro
        raise NotImplementedError


class ParameterNode(NodeBase):
    """
    Represents a <Parameter> node.
    """

    def to_excel_row(self) -> dict:
        row = {"TagType": "Parameter", "Name": "", "FullPath": self.fullpath}
        # Name
        row["Name"] = self.original_subs.get("Name", "")
        # Standard fields
        for col in (
            "Real",
            "Integer",
            "High",
            "Low",
            "String",
            "EnumerationSet",
            "EnumerationMember",
        ):
            row[col] = self.original_subs.get(col, "")
        # Parameters never have Defer
        row["Defer"] = ""
        # No FormulaValueLimit for Parameter
        # Capture any "extras" beyond EXCEL_COLUMNS
        for k, v in self.original_subs.items():
            if k not in row:
                row[k] = v
        return row


class FormulaValueNode(NodeBase):
    """
    Represents a <FormulaValue> node.
    """

    def to_excel_row(self) -> dict:
        row = {"TagType": "FormulaValue", "Name": "", "FullPath": self.fullpath}
        row["Name"] = self.original_subs.get("Name", "")
        # Value vs Defer
        row["Defer"] = self.original_subs.get("Defer", "")
        if row["Defer"] == "":
            # non-deferred => blank defer, but Value tag exists
            row["Value"] = self.original_subs.get("Value", "")
        else:
            row["Value"] = ""
        # Type fields
        for col in ("Real", "Integer"):
            row[col] = self.original_subs.get(col, "")
        for col in ("String", "EnumerationSet", "EnumerationMember"):
            row[col] = self.original_subs.get(col, "")
        # FVL fields
        # Verification attribute
        fvl = self.element.find(f"{{{NAMESPACE}}}FormulaValueLimit", namespaces=NSMAP)
        if fvl is not None:
            row["FormulaValueLimit_Verification"] = fvl.get("Verification", "")
            for child in fvl:
                name = QName(child.tag).localname
                row[f"FormulaValueLimit_{name}"] = child.text or ""
        else:
            # blank out FVL columns
            for col in EXCEL_COLUMNS:
                if col.startswith("FormulaValueLimit_"):
                    row[col] = ""
        # extras
        for k, v in self.original_subs.items():
            if k not in row:
                row[k] = v
        return row


class RecipeTree:
    """
    Holds XML tree plus extracted ParameterNode & FormulaValueNode lists.
    """

    def __init__(self, path: str):
        self.filepath = path
        self.tree = etree.parse(path)
        self.root = self.tree.getroot()
        self.parameters = []  # List[ParameterNode]
        self.formula_values = []  # List[FormulaValueNode]

    def extract_nodes(self):
        """
        Populate self.parameters & self.formula_values by walking the XML.
        """
        rid_el = self.root.find(f"{{{NAMESPACE}}}RecipeElementID")
        recipe_id = (rid_el.text or "") if rid_el is not None else ""
        # Parameters (direct children)
        for p in self.root.findall(f"{{{NAMESPACE}}}Parameter"):
            name = p.find(f"{{{NAMESPACE}}}Name").text or ""
            fp = f"{recipe_id}/Parameter[{name}]"
            self.parameters.append(ParameterNode(p, fp, self.filepath))

        # FormulaValues under Steps/Step
        for fv in self.root.findall(f".//{{{NAMESPACE}}}FormulaValue"):
            # locate parent <Step> with Name
            step = fv.getparent()
            step_name = step.find(f"{{{NAMESPACE}}}Name").text or ""
            name = fv.find(f"{{{NAMESPACE}}}Name").text or ""
            fp = f"{recipe_id}/Steps/Step[{step_name}]/FormulaValue[{name}]"
            self.formula_values.append(FormulaValueNode(fv, fp, self.filepath))
