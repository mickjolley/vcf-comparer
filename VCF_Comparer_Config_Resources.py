# -*- coding: utf-8 -*-
"""
VP_Config_Resources.py - Resource definitions for VCF Comparer Config GUI

This module contains all static resource definitions for the VP Config GUI,
following boa-constructor methodology for modular resource management.

Boa Constructor Pattern:
- Separates resource definitions from UI logic
- Centralizes configuration tooltips, labels, and field definitions
- Enables easy localization and maintenance
"""

# Field definitions with metadata
FIELD_DEFINITIONS = {
    
    # VCF FilePath Fields
    'VCF_FILE_PATH': {
        'type': 'path_directory',
        'label': 'VCF File Path',
        'tooltip': 'Path to VCF file.',
        'required': True,
    },

    # Path Fields
    'DNA_FILES_PATH': {
        'type': 'path_directory',
        'label': 'DNA Files Path',
        'tooltip': 'Path to folder where the DNA files are stored.',
        'required': True,
    },
    'WORKING_DIRECTORY': {
        'type': 'path_directory',
        'label': 'Working Directory',
        'tooltip': 'Folder where the .xlsx and .py files will be stored.',
        'required': True,
    },
    'MAP_PATH': {
        'type': 'path_directory',
        'label': 'Map File Path (min_map.txt)',
        'tooltip': 'Path to folder containing min_map.txt.',
        'required': True,
    },

    # File Lists
    'INDIVIDUALS': {
        'type': 'list_text',
        'label': 'VCF Individuals (comma-separated names)',
        'tooltip': "Comma-separated names of the individuals in th VCF file to compare (minimum 2). Enter '*' for all",
        'required': True,
    },
    'SUBJECTS': {
        'type': 'list_files',
        'label': 'Subjects (comma-separated names)',
        'tooltip': "Comma-separated names of subjects in the DNA file to compare with each other or individuals in the VCF File. Enter '*' for all.",
        'required': False,
    },
    'CHROMOSOMES': {
        'type': 'list_text',
        'label': 'Chromosomes (1-22, comma-separated or empty for all)',
        'tooltip': 'Chromosome selected (1-23). Leave empty for all chromosomes.',
        'required': False,
    },

    # Text Fields
    'EXCEL_FILE_NAME': {
        'type': 'text',
        'label': 'Excel Output File Name',
        'tooltip': 'Name of the .xlsx file. Do not include the ".xlsx".',
        'required': True,
    },
    'LINUX_FONT_STRING': {
        'type': 'text',
        'label': 'Linux Font Path',
        'tooltip': 'Linux users only. Enter the path to your font.',
        'required': True,
        'default': '*/fonts/truetype/family/DejaVuSerif-Bold.ttf',
    },
    'FREEZE_COLUMN': {
        'type': 'text',
        'label': 'Freeze Column (e.g., "A" for no freeze)',
        'tooltip': 'Set to "A" if freezing not desired.',
        'required': True,
        'default': "A",
    },

    # Boolean Fields
    'SHOW_NO_MATCHES': {
        'type': 'boolean',
        'label': 'SHOW_NO_MATCHES',
        'tooltip': 'Set to False if display of match pairs with no matching segments is not desired.',
        'default': True,
    },
    'CHROM_TRUE_SIZE': {
        'type': 'boolean',
        'label': 'CHROM_TRUE_SIZE',
        'tooltip': 'Set to True for true size. Set to False for normalized size.',
        'default': False,
    },
    'LINEAR_CHROMOSOME': {
        'type': 'boolean',
        'label': 'LINEAR_CHROMOSOME',
        'tooltip': 'Set to True if you want to see the linearized chromosomes.',
        'default': False,
    },
    'PARENTAL_RELATIONSHIP': {
        'type': 'boolean',
        'label': 'PARENTAL_rELATIONSHIP',
        'tooltip': 'Set to True if the calculation of ROHs in subjects is desired.',
        'default': False,
    },
    'FIR_TABLES': {
        'type': 'boolean',
        'label': 'FIR_TABLES',
        'tooltip': 'Set to True if display of FIR tables is desired.',
        'default': True,
    },
    'SCALE_ON': {
        'type': 'boolean',
        'label': 'SCALE_ON',
        'tooltip': 'Turn scale on and off.',
        'default': True,
    },

    # Numeric Integer Fields
    'RESOLUTION': {
        'type': 'int_spin',
        'label': 'Resolution',
        'tooltip': 'Default value = 1. Keep under 10. Set to 100 for full resolution.',
        'min': 1,
        'max': 100,
        'default': 1,
    },
    'HIR_SNP_MIN': {
        'type': 'int_spin',
        'label': 'HIR SNP Minimum',
        'tooltip': 'Minimum number of HIR SNPs. Default value = 200.',
        'min': 1,
        'max': 10000,
        'default': 200,
    },
    'FIR_SNP_MIN': {
        'type': 'int_spin',
        'label': 'FIR SNP Minimum',
        'tooltip': 'Minimum number of FIR SNPs. Default value = 75.',
        'min': 1,
        'max': 10000,
        'default': 75,
    },
    'MM_DIST': {
        'type': 'int_spin',
        'label': 'Mismatch Distance (Kbs)',
        'tooltip': 'Number of Kbs between mismatches to end segment. Default = 1000.',
        'min': 1,
        'max': 10000,
        'default': 1000,
    },
    'PC_NO_CALLS_ALLOWED': {
        'type': 'int_spin',
        'label': 'Percent No Calls Allowed',
        'tooltip': 'Percentage no-calls allowed in files.',
        'min': 0,
        'max': 100,
        'default': 10,
    },
    # Float Fields
    'SCALE_FACTOR': {
        'type': 'float',
        'label': 'Scale Factor',
        'tooltip': 'The column width per pixel factor. Default value = 0.1351.',
        'default': 0.1355,
    },
    'HIR_CUTOFF': {
        'type': 'float',
        'label': 'HIR Cutoff (cM)',
        'tooltip': 'HIR Minimum segment length (cM). The default is 7.',
        'min': 0,
        'max': 100,
        'default': 7,
    },
    'FIR_CUTOFF': {
        'type': 'float',
        'label': 'FIR Cutoff (cM)',
        'tooltip': 'FIR cutoff. FIRs less than 1cM in length are probably not significant.',
        'min': 0,
        'max': 100,
        'default': 1,
    },
}

# Section groupings for UI organization
FIELD_SECTIONS = {
    'Paths and Directories': [
        'VCF_FILE_PATH',
        'DNA_FILES_PATH',
        'WORKING_DIRECTORY',
        'MAP_PATH',
    ],
    'File Management': [
        'INDIVIDUALS',
        'SUBJECTS',
        'CHROMOSOMES',
        'EXCEL_FILE_NAME',
    ],
    'Display and Processing Options': [
        'SHOW_NO_MATCHES',
        'CHROM_TRUE_SIZE',
        'LINEAR_CHROMOSOME',
        'FIR_TABLES',
        'SCALE_ON',
        'PC_NO_CALLS_ALLOWED',
        'PARENTAL_RELATIONSHIP',
    ],
    'Numeric Parameters': [
        'RESOLUTION',
        'HIR_CUTOFF',
        'FIR_CUTOFF',
        'HIR_SNP_MIN',
        'FIR_SNP_MIN',
        'MM_DIST',
        'SCALE_FACTOR',
    ],
    'Other Settings': [
        'FREEZE_COLUMN',
        'LINUX_FONT_STRING',
    ],
}

# Application strings
APP_TITLE = 'VCF Comparer Configuration Editor'
APP_VERSION = '1.0'
APP_DESCRIPTION = 'GUI Configuration Editor for VCF_Comparer_configV3.py'
APP_AUTHOR = 'Mick Jolley'
APP_EMAIL = 'mickj1948@gmail.com'
APP_COPYRIGHT = '© 2026 Mick Jolley'
APP_WEBSITE = 'https://github.com/mickjolley/vcf-comparer'

# Messages
MSG_LOAD_SUCCESS = 'Configuration loaded successfully'
MSG_SAVE_SUCCESS = 'Configuration saved successfully!'
MSG_LOAD_ERROR = 'Error loading configuration: {}'
MSG_SAVE_ERROR = 'Error saving configuration: {}'
MSG_RESET_CONFIRM = 'Reset all values to defaults?'
MSG_INVALID_FLOAT = 'Invalid float value for {}: {}'

# UI Constants
WINDOW_WIDTH = 900
WINDOW_HEIGHT = 800
CONTROL_HEIGHT = 60
PADDING = 5
SECTION_FONT_SIZE = 12
TITLE_FONT_SIZE = 14

# File dialog wildcards
WILDCARD_PYTHON = 'Python files (*.py)|*.py|All files (*.*)|*.*'
WILDCARD_ALL = 'All files (*.*)|*.*'
WILDCARD_TXT = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'


def get_field_definition(field_name):
    """Get field definition by name"""
    return FIELD_DEFINITIONS.get(field_name, {})


def get_field_section(field_name):
    """Get the section that contains the field"""
    for section, fields in FIELD_SECTIONS.items():
        if field_name in fields:
            return section
    return 'Other Settings'


def get_all_sections():
    """Get all section names in order"""
    return list(FIELD_SECTIONS.keys())


def get_fields_in_section(section_name):
    """Get all fields in a section"""
    return FIELD_SECTIONS.get(section_name, [])
