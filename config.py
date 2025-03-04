import json
import os
import sys


def get_config_path(f_name='config.json'):
    """Возвращает правильный путь к config.json в зависимости от режима работы (обычный или exe)."""
    if getattr(sys, 'frozen', False):  # Если программа запущена как .exe
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  # Папка с этим скриптом

    return os.path.join(base_path, f_name)

config_path = get_config_path()

with open(config_path, "r", encoding='utf-8') as f:
    config = json.load(f)


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
TYPE_ATTR_NAME = config['TYPE_ATTR_NAME']
TASK_TYPE_ATTR_NAME = config['TASK_TYPE_ATTR_NAME']
IFC_ATTR_NAME = config['IFC_ATTR_NAME']
ASM_TYPE_ATTR_NAME = config['ASM_TYPE_ATTR_NAME']
SPECIALITY_ATTR_NAME = config["SPECIALITY_ATTR_NAME"]
ASSEMBLY_TYPE_ATTR_NAME = config['ASSEMBLY_TYPE_ATTR_NAME']
TAG_ATTR_NAME = config["TAG_ATTR_NAME"]
NAME_ATTR_NAME = config["NAME_ATTR_NAME"]
ASSEMBLY_MARK_ATTR_NAME = config['ASSEMBLY_MARK_ATTR_NAME']
ASSEMBLY_N_ATTR_NAME = config["ASSEMBLY_N_ATTR_NAME"]
PROJECT_MARK_ATTR_NAME = config["PROJECT_MARK_ATTR_NAME"]
SYSTEM_TAG_ATTR_NAME = config["SYSTEM_TAG_ATTR_NAME"]
EL_LINE_TAG_ATTR_NAME = config["EL_LINE_TAG_ATTR_NAME"]
TASK_AUTHOR_ATTR_NAME = config["TASK_AUTHOR_ATTR_NAME"]
TASK_USER_ATTR_NAME = config["TASK_USER_ATTR_NAME"]
SPECIALITIES = config['SPECIALITIES']

"""report_profile"""
# REPORT_TYPES - набор для формирования блока Types в xml-схеме отчета для CadLib
REPORT_TYPES = config['REPORT_TYPES']

"""d_parser"""
SHEET_NAMES = config['SHEET_NAMES']
