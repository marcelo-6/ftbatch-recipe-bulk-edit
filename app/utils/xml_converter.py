# utils/xml_converter.py
from utils.string_converter import camel_to_snake, snake_to_camel, snake_to_pascal
from lxml import etree
from models.recipe import (
    RecipeElement, Header, Parameter, Steps, Step, FormulaValue,
    PhaseLinkGroup, PhaseLink, Transition, ElementLink,
    UnitRequirement, DownstreamResource, Formulations, Formulation, FormulationParameter
)
from datetime import datetime
from typing import Optional

class XMLConverter:
    def __init__(self, xsd_path: Optional[str] = None):
        self.xsd_path = xsd_path
        if xsd_path:
            with open(xsd_path, 'rb') as f:
                schema_doc = etree.XML(f.read())
                self.schema = etree.XMLSchema(schema_doc)
        else:
            self.schema = None

    def parse_xml(self, xml_file: str) -> RecipeElement:
        tree = etree.parse(xml_file)
        root = tree.getroot()
        if self.schema:
            self.schema.assertValid(tree)
        NS = "urn:Rockwell/MasterRecipe"
        
        # Root attribute
        schema_version = root.get("SchemaVersion")
        
        # RecipeElementID
        recipe_element_id = root.findtext(f"{{{NS}}}RecipeElementID")
        
        # Header
        header_el = root.find(f"{{{NS}}}Header")
        header_dict = {}
        if header_el is not None:
            for child in header_el:
                tag = etree.QName(child).localname
                header_dict[tag.lower()] = child.text
        header = Header(**header_dict)
        
        # Parameters
        parameters = []
        for param in root.findall(f"{{{NS}}}Parameter"):
            param_dict = {}
            for child in param:
                tag = etree.QName(child).localname
                param_dict[tag.lower()] = child.text
            parameters.append(Parameter(**param_dict))
        
        # Steps
        steps_el = root.find(f"{{{NS}}}Steps")
        steps_obj = Steps()
        if steps_el is not None:
            for child in steps_el:
                tag = etree.QName(child).localname.lower()
                step_dict = {}
                # Parse attributes
                x_pos_attr = child.get("XPos")
                y_pos_attr = child.get("YPos")
                if x_pos_attr is not None:
                    step_dict["x_pos"] = int(x_pos_attr)
                if y_pos_attr is not None:
                    step_dict["y_pos"] = int(y_pos_attr)
                acquire_unit = child.get("AcquireUnit")
                system_step = child.get("SystemStep")
                if acquire_unit is not None:
                    step_dict["acquire_unit"] = acquire_unit.lower() == "true"
                if system_step is not None:
                    step_dict["system_step"] = system_step.lower() == "true"
                # Parse child elements
                for sub in child:
                    subtag = etree.QName(sub).localname
                    if subtag == "FormulaValue":
                        fv_dict = {}
                        for fv_child in sub:
                            fv_tag = etree.QName(fv_child).localname
                            fv_dict[fv_tag.lower()] = fv_child.text
                        # Ensure we have a list to collect formula values
                        step_dict.setdefault("formula_values", []).append(FormulaValue(**fv_dict))
                    else:
                        # Convert element names to lower-case to match our snake_case fields
                        step_dict[subtag.lower()] = sub.text
                # Convert numeric fields if necessary
                if "packedflags" in step_dict and step_dict["packedflags"]:
                    step_dict["packed_flags"] = int(step_dict.pop("packedflags"))
                step_obj = Step(**step_dict)
                if tag == "initialstep":
                    steps_obj.initial_step = step_obj
                elif tag == "terminalstep":
                    steps_obj.terminal_step = step_obj
                elif tag == "step":
                    steps_obj.steps.append(step_obj)
        else:
            steps_obj = None

        # PhaseLinkGroup
        plg_el = root.find(f"{{{NS}}}PhaseLinkGroup")
        phase_link_group = None
        if plg_el is not None:
            plg_dict = {}
            plg_dict["name"] = plg_el.findtext(f"{{{NS}}}Name")
            phase_links = []
            for pl in plg_el.findall(f"{{{NS}}}PhaseLink"):
                pl_dict = {}
                for child in pl:
                    tag = etree.QName(child).localname
                    tag = camel_to_snake(tag).lower()
                    pl_dict[tag] = child.text
                phase_links.append(PhaseLink(**pl_dict))
            plg_dict["phase_links"] = phase_links
            phase_link_group = PhaseLinkGroup(**plg_dict)
        
        # Transitions
        transitions = []
        for trans in root.findall(f"{{{NS}}}Transition"):
            trans_dict = {}
            trans_dict["name"] = trans.findtext(f"{{{NS}}}Name")
            trans_dict["conditional_expression"] = trans.findtext(f"{{{NS}}}ConditionalExpression")
            trans_dict["x_pos"] = trans.get("XPos")
            trans_dict["y_pos"] = trans.get("YPos")
            transitions.append(Transition(**trans_dict))
        
        # ElementLinks
        element_links = []
        for elink in root.findall(f"{{{NS}}}ElementLink"):
            elink_dict = {}
            for child in elink:
                tag = etree.QName(child).localname
                elink_dict[tag.lower()] = child.text
            element_links.append(ElementLink(**elink_dict))
        
        # UnitRequirements
        unit_requirements = []
        for ur in root.findall(f"{{{NS}}}UnitRequirement"):
            ur_dict = {}
            for child in ur:
                tag = etree.QName(child).localname
                if tag == "DownstreamResource":
                    dr_dict = {}
                    for dr_child in child:
                        dr_tag = etree.QName(dr_child).localname
                        dr_tag = camel_to_snake(dr_tag).lower()
                        dr_dict[dr_tag] = dr_child.text
                    ur_dict["downstream_resource"] = DownstreamResource(**dr_dict)
                else:
                    tag = camel_to_snake(tag).lower()
                    ur_dict[tag] = child.text
            unit_requirements.append(UnitRequirement(**ur_dict))
        
        # Comments
        comments_el = root.find(f"{{{NS}}}Comments")
        comments = comments_el.text if comments_el is not None else None
        
        # Formulations
        formulations_el = root.find(f"{{{NS}}}Formulations")
        formulations = None
        if formulations_el is not None:
            formulations_list = []
            for form_el in formulations_el.findall(f"{{{NS}}}Formulation"):
                form_dict = {}
                form_dict["name"] = form_el.findtext(f"{{{NS}}}Name")
                form_dict["description"] = form_el.findtext(f"{{{NS}}}Description")
                params_list = []
                param_list_el = form_el.find(f"{{{NS}}}ParameterList")
                if param_list_el is not None:
                    for p in param_list_el.findall(f"{{{NS}}}Parameter"):
                        p_dict = {}
                        for child in p:
                            tag = etree.QName(child).localname
                            p_dict[tag.lower()] = child.text
                        params_list.append(FormulationParameter(**p_dict))
                form_dict["parameter_list"] = params_list
                formulations_list.append(Formulation(**form_dict))
            formulations = Formulations(formulations=formulations_list)
        
        recipe = RecipeElement(
            schema_version=schema_version,
            recipe_element_id=recipe_element_id,
            header=header,
            parameters=parameters,
            steps=steps_obj,
            phase_link_group=phase_link_group,
            transitions=transitions,
            element_links=element_links,
            unit_requirements=unit_requirements,
            comments=comments,
            formulations=formulations
        )
        return recipe

    def to_xml(self, recipe: RecipeElement) -> etree._Element:
        NS = "urn:Rockwell/MasterRecipe"
        nsmap = {None: NS, "xsi": "http://www.w3.org/2001/XMLSchema-instance"}
        root = etree.Element("RecipeElement", nsmap=nsmap)
        root.set("SchemaVersion", recipe.schema_version or "3530")
        
        # RecipeElementID
        id_el = etree.SubElement(root, etree.QName(NS, "RecipeElementID"))
        id_el.text = recipe.recipe_element_id
        
        # Header
        header_el = etree.SubElement(root, etree.QName(NS, "Header"))
        for field, value in recipe.header.dict().items():
            child = etree.SubElement(header_el, etree.QName(NS, snake_to_pascal(field)))#.capitalize()))
            child.text = str(value) if value is not None else ""
        
        # Parameters
        for param in recipe.parameters:
            param_el = etree.SubElement(root, etree.QName(NS, "Parameter"))
            for field, value in param.dict().items():
                child = etree.SubElement(param_el, etree.QName(NS, snake_to_pascal(field)))
                child.text = str(value) if value is not None else ""
        
        # Steps
        if recipe.steps:
            steps_el = etree.SubElement(root, etree.QName(NS, "Steps"))
            if recipe.steps.initial_step:
                init_el = self._step_to_xml(recipe.steps.initial_step, NS, "InitialStep")
                steps_el.append(init_el)
            if recipe.steps.terminal_step:
                term_el = self._step_to_xml(recipe.steps.terminal_step, NS, "TerminalStep")
                steps_el.append(term_el)
            for step in recipe.steps.steps:
                step_el = self._step_to_xml(step, NS, "Step")
                steps_el.append(step_el)
        
        # PhaseLinkGroup
        if recipe.phase_link_group:
            plg_el = etree.SubElement(root, etree.QName(NS, "PhaseLinkGroup"))
            name_el = etree.SubElement(plg_el, etree.QName(NS, "Name"))
            name_el.text = recipe.phase_link_group.name
            for pl in recipe.phase_link_group.phase_links:
                pl_el = etree.SubElement(plg_el, etree.QName(NS, "PhaseLink"))
                for field, value in pl.dict().items():
                    child = etree.SubElement(pl_el, etree.QName(NS, snake_to_pascal(field)))
                    child.text = str(value) if value is not None else ""
        
        # Transitions
        for trans in recipe.transitions:
            trans_el = etree.SubElement(root, etree.QName(NS, "Transition"))
            trans_el.set("XPos", str(trans.x_pos) if trans.x_pos is not None else "0")
            trans_el.set("YPos", str(trans.y_pos) if trans.y_pos is not None else "0")
            name_el = etree.SubElement(trans_el, etree.QName(NS, "Name"))
            name_el.text = trans.name
            cond_el = etree.SubElement(trans_el, etree.QName(NS, "ConditionalExpression"))
            cond_el.text = trans.conditional_expression if trans.conditional_expression is not None else ""
        
        # ElementLinks
        for elink in recipe.element_links:
            elink_el = etree.SubElement(root, etree.QName(NS, "ElementLink"))
            for field, value in elink.dict().items():
                child = etree.SubElement(elink_el, etree.QName(NS, snake_to_pascal(field)))
                child.text = str(value) if value is not None else ""
        
        # UnitRequirements
        for ur in recipe.unit_requirements:
            ur_el = etree.SubElement(root, etree.QName(NS, "UnitRequirement"))
            for field, value in ur.dict().items():
                if field == "downstream_resource" and value:
                    dr_el = etree.SubElement(ur_el, etree.QName(NS, "DownstreamResource"))
                    for dr_field, dr_value in value.items():
                        child = etree.SubElement(dr_el, etree.QName(NS, snake_to_pascal(dr_field)))
                        child.text = str(dr_value) if dr_value is not None else ""
                else:
                    child = etree.SubElement(ur_el, etree.QName(NS, snake_to_pascal(field)))
                    child.text = str(value) if value is not None else ""
        
        # Comments
        comments_el = etree.SubElement(root, etree.QName(NS, "Comments"))
        comments_el.text = recipe.comments if recipe.comments is not None else ""
        
        # Formulations
        if recipe.formulations:
            formulations_el = etree.SubElement(root, etree.QName(NS, "Formulations"))
            for form in recipe.formulations.formulations:
                form_el = etree.SubElement(formulations_el, etree.QName(NS, "Formulation"))
                name_el = etree.SubElement(form_el, etree.QName(NS, "Name"))
                name_el.text = form.name
                desc_el = etree.SubElement(form_el, etree.QName(NS, "Description"))
                desc_el.text = form.description if form.description is not None else ""
                param_list_el = etree.SubElement(form_el, etree.QName(NS, "ParameterList"))
                for fp in form.parameter_list:
                    fp_el = etree.SubElement(param_list_el, etree.QName(NS, "Parameter"))
                    for field, value in fp.dict().items():
                        child = etree.SubElement(fp_el, etree.QName(NS, snake_to_pascal(field)))
                        child.text = str(value) if value is not None else ""
        
        return root

    def _step_to_xml(self, step: Step, NS: str, tag: str) -> etree._Element:
        el = etree.Element(etree.QName(NS, tag))
        if step.x_pos is not None:
            el.set("XPos", str(step.x_pos))
        if step.y_pos is not None:
            el.set("YPos", str(step.y_pos))
        if step.acquire_unit is not None:
            el.set("AcquireUnit", "true" if step.acquire_unit else "false")
        if step.system_step is not None:
            el.set("SystemStep", "true" if step.system_step else "false")
        for key in ["step_recipe_id", "packed_flags", "unit_alias"]:
            value = getattr(step, key)
            if value is not None:
                child = etree.SubElement(el, etree.QName(NS, snake_to_pascal(key)))
                child.text = str(value)
        name_el = etree.SubElement(el, etree.QName(NS, "Name"))
        name_el.text = step.name
        for fv in step.formula_values:
            fv_el = etree.SubElement(el, etree.QName(NS, "FormulaValue"))
            for field, value in fv.dict().items():
                child = etree.SubElement(fv_el, etree.QName(NS, snake_to_pascal(field)))
                child.text = str(value) if value is not None else ""
        return el
