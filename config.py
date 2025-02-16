# for WSHT2024
# TYPE_ATTR_NAME = 'RD_TYPE'

# default
# TYPE_ATTR_NAME = 'TYPE'
# TASK_TYPE_ATTR_NAME = 'TASK_TYPE'
# IFC_ATTR_NAME = 'IFC_TYPE'
# ASM_TYPE_ATTR_NAME = 'ASSEMBLY_TYPE'
# SPECIALITY_ATTR_NAME = "SPECIALITY"
# ASSEMBLY_TYPE_ATTR_NAME = 'ASSEMBLY_TYPE'
# TAG_ATTR_NAME = "TAG"
# NAME_ATTR_NAME = "NAME"
# ASSEMBLY_MARK_ATTR_NAME = 'ASSEMBLY_MARK'
# ASSEMBLY_N_ATTR_NAME = "ASSEMBLY_N"
# PROJECT_MARK_ATTR_NAME = "PROJECT_MARK"
# SYSTEM_TAG_ATTR_NAME = "SYSTEM_TAG"
# EL_LINE_TAG_ATTR_NAME = "EL_LINE_TAG"
# TASK_AUTHOR_ATTR_NAME ="TASK_AUTHOR"
# TASK_USER_ATTR_NAME = "TASK_USER"

# YS2025
TYPE_ATTR_NAME = 'ED_TYPE'
TASK_TYPE_ATTR_NAME = 'ED_TASK_TYPE'
IFC_ATTR_NAME = 'IFC_TYPE'
ASM_TYPE_ATTR_NAME = 'ED_ASSEMBLY_TYPE'
SPECIALITY_ATTR_NAME = "ED_SPECIALITY"
ASSEMBLY_TYPE_ATTR_NAME = 'ED_ASSEMBLY_TYPE'
TAG_ATTR_NAME = "ED_TAG"
NAME_ATTR_NAME = "ED_NAME"
ASSEMBLY_MARK_ATTR_NAME = 'ED_ASSEMBLY_MARK'
ASSEMBLY_N_ATTR_NAME = "ED_ASSEMBLY_N"
PROJECT_MARK_ATTR_NAME = "ED_PROJECT_MARK"
SYSTEM_TAG_ATTR_NAME = "ED_SYSTEM_TAG"
EL_LINE_TAG_ATTR_NAME = "ED_EL_LINE_TAG"
TASK_AUTHOR_ATTR_NAME ="ED_TASK_AUTHOR"
TASK_USER_ATTR_NAME = "ED_TASK_AUTHOR"

"""report_profile"""
# REPORT_TYPES - набор для формирования блока Types в xml-схеме отчета для CadLib
REPORT_TYPES = ['AEC_SURF', 'COLLISIONS', 'Cable_ETS', 'HVAC', 'HVACAxis', 'HVACPipeItems', 'IFCElement', 'Insulation',
         'Materials', 'OptionsList', 'REINFORCEMENT', 'TRAYS', 'TowerSettings', 'WireDefault', 'WorkItem', 'WorkList',
         'WorkResource', 'aec_elements', 'aec_layer_groups', 'aec_layer_material', 'assemblies', 'assembly', 'bolts',
         'cabinet', 'cables', 'cat160', 'cat161', 'cat47', 'catDiameterList', 'catLineSystems', 'catMaterial',
         'catMaterials', 'catPipeSupportStepData', 'cat_clash', 'climate', 'commonLinks', 'concreteware', 'decoration',
         'device', 'foundation', 'garlandMisc', 'garlands', 'gesnCategoryList', 'gesnResourceItem', 'ground',
         'hierarchical_list', 'iron_assemblies', 'label', 'metalware', 'metalwarenode', 'modifiers', 'node', 'nodes',
         'panels', 'pipeAxis', 'pipeCategoryList', 'pipeItems', 'pipeLineNumberList', 'pipeMaterialList', 'report',
         'routeConstructions', 'routeConstructionsAssemblies', 'structure_data', 'suppressor', 'symbol', 'tank_eq_schm',
         'tank_equip', 'towerequipment', 'units', 'valve_prototype', 'wire', 'wires', 'zoneList']

"""d_parser"""
SHEET_NAMES = ['AEC_EL', 'PIPE_EL', 'CABLE_EL', 'TASK']

"""ifc_import"""
PROPERTY_SET_NAME = 'EngeneeringDesign'
