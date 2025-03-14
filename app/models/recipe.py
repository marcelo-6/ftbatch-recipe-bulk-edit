# models/recipe.py
from typing import List, Optional
from pydantic import BaseModel, Field
from datetime import datetime

# ---------------------------
# Header Section
# ---------------------------
class Header(BaseModel):
    abstract: Optional[str] = None
    approved_by: Optional[str] = None
    area_model: Optional[str] = None
    area_model_date: Optional[datetime] = None
    author: Optional[str] = None
    being_edited_by: Optional[str] = None
    class_based: Optional[bool] = None
    db_schema: Optional[str] = None
    default_size: Optional[float] = None
    description: Optional[str] = None
    duration: Optional[int] = None
    locale_id: Optional[int] = None
    max_size: Optional[float] = None
    min_size: Optional[float] = None
    product_code: Optional[str] = None
    product_id: Optional[str] = None
    product_units: Optional[str] = None
    recipe_type: Optional[str] = None
    released: Optional[int] = None
    release_as_step: Optional[bool] = None
    resource: Optional[str] = None
    verification_date: Optional[datetime] = None
    version: Optional[str] = None
    version_date: Optional[datetime] = None
    obsoleted: Optional[bool] = None
    next_wip_number: Optional[int] = None
    version_description: Optional[str] = None
    parent_name: Optional[str] = None
    parent_version_description: Optional[str] = None
    parent_version_date: Optional[datetime] = None
    parent_verification_date: Optional[datetime] = None
    parent_area_model_date: Optional[datetime] = None
    parent_area_model_name: Optional[str] = None
    security_authority_identifier: Optional[str] = None

# ---------------------------
# Parameter Section
# ---------------------------
class Parameter(BaseModel):
    name: str
    erp_alias: Optional[str] = None
    plc_reference: Optional[int] = None
    real: Optional[float] = None
    high: Optional[float] = None
    low: Optional[float] = None
    engineering_units: Optional[str] = None
    scale: Optional[bool] = None

# ---------------------------
# FormulaValue Section (within Steps)
# ---------------------------
class FormulaValue(BaseModel):
    name: str
    display: Optional[bool] = None
    defer: Optional[str] = None
    real: Optional[float] = None
    engineering_units: Optional[str] = None
    value: Optional[str] = None
    string: Optional[str] = None

# ---------------------------
# Step Section
# ---------------------------
class Step(BaseModel):
    name: str
    x_pos: Optional[int] = None
    y_pos: Optional[int] = None
    acquire_unit: Optional[bool] = None
    system_step: Optional[bool] = None
    step_recipe_id: Optional[str] = None
    packed_flags: Optional[int] = None
    unit_alias: Optional[str] = None
    formula_values: List[FormulaValue] = Field(default_factory=list)

class Steps(BaseModel):
    initial_step: Optional[Step] = None
    terminal_step: Optional[Step] = None
    steps: List[Step] = Field(default_factory=list)

# ---------------------------
# PhaseLinkGroup Section
# ---------------------------
class PhaseLink(BaseModel):
    unit_procedure_step_name: str
    operation_step_name: str
    phase_step_name: str

class PhaseLinkGroup(BaseModel):
    name: str
    phase_links: List[PhaseLink] = Field(default_factory=list)

# ---------------------------
# Transition Section
# ---------------------------
class Transition(BaseModel):
    name: str
    x_pos: Optional[int] = None
    y_pos: Optional[int] = None
    conditional_expression: Optional[str] = None

# ---------------------------
# ElementLink Section
# ---------------------------
class ElementLink(BaseModel):
    from_transition: Optional[str] = None
    to_step: Optional[str] = None
    from_step: Optional[str] = None
    to_transition: Optional[str] = None

# ---------------------------
# UnitRequirement Section
# ---------------------------
class DownstreamResource(BaseModel):
    name: str

class UnitRequirement(BaseModel):
    unit_alias: str
    class_instance: str
    binding_method: str
    material_binding_method: str
    class_based: Optional[bool] = None
    downstream_resource: Optional[DownstreamResource] = None

# ---------------------------
# Formulations Section
# ---------------------------
class FormulationParameter(BaseModel):
    name: str
    real: Optional[float] = None

class Formulation(BaseModel):
    name: str
    description: Optional[str] = None
    parameter_list: List[FormulationParameter] = Field(default_factory=list)

class Formulations(BaseModel):
    formulations: List[Formulation] = Field(default_factory=list)

# ---------------------------
# Root Element
# ---------------------------
class RecipeElement(BaseModel):
    schema_version: Optional[str] = None
    recipe_element_id: str
    header: Header
    parameters: List[Parameter] = Field(default_factory=list)
    steps: Optional[Steps] = None
    phase_link_group: Optional[PhaseLinkGroup] = None
    transitions: List[Transition] = Field(default_factory=list)
    element_links: List[ElementLink] = Field(default_factory=list)
    unit_requirements: List[UnitRequirement] = Field(default_factory=list)
    comments: Optional[str] = None
    formulations: Optional[Formulations] = None
