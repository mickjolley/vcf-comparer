# -*- coding: utf-8 -*-
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
VCF_Comparer_configV3.py is the configuration file for VCF_Comparer.V3.0.py.

VCF_FILE_PATH: Path to .vcf file. Include ".vcf" suffix.

DNA_FILES_PATH: Path to DNA files.

WORKING_DIRECTORY: Folder where the .xlsx and .csv files will be stored.

MAP_PATH: Path to folder containing min_map.txt.

INDIVIDUALS: List of individuals to load from the VCF file. Enter ['*'] to load
all individuals. Leave empty ([]) to load none.

SUBJECTS: List of individuals to be compared from DNA_FILES_PATH. Enter ['*'] 
to load all subjects. Leave empty ([]) to load none. SUBJECT file names should 
contain '_raw_dna' after the name.

CHROMOSOMES: Chromosome selected (1-23). More than one chromosome may be entered.
Leave empty for all chromosomes.

EXCEL_FILE_NAME: Name of the .xlsx file. Do not include the ".xlsx", This is
added automatically.

SHOW_NO_MATCHES: Set to False if the display of match pairs with no matching
segments is not desired. This is the recommended default value for cousin
matches. If only siblings are being compared set this to True.

CHROM_TRUE_SIZE: Set to True for true size. Set to False for normalized size.

LINEAR_CHROMOSOME: Set to True if you want to see the linearized chromosomes.
RESOLUTION will be ignored unless it is 10 (10x resolution). CHROM_TRUE_SIZE is
automatically set to False.

RESOLUTION: Default value = 1. For normalized size it is advised to keep it
under 10. Set to 100 for full length chromosomes. If LINEAR_CHROMOSOME is set
to "True", RESOLUTION is automatically set to 1, unless it is set to 10.

HIR_CUTOFF: Default value = 7 cM

FIR_CUTOFF: Default value = 1 cM.

X_HIR_CUTOFF: X chromosome cutoff (cM). The default is 15.

X_FIR_CUTOFF: X chromosome FIR cutoff (cM). The default 15.

FIR_TABLES: Set to True if display of FIR tables is desired.

SCALE_ON: Turn scale on and off. Set to False if not required. Default = True

FREEZE_COLUMN: Set to "A" if freezing not desired. Default = "A".

LINUX_FONT_STRING: Linux users only. Enter the path to your font. If you don't
know it, set SCALE_ON to False.

SHOW_TIMES: Elapsed times are shown for each step. Default =True

SHOW_MATCH_PAIR_PROGRESS: Notifies the completion of each step. Set to
False if you don't want to see this. Default = True

HIR_SNP_MIN: Minimum number of HIR SNPs. Default value = 200

FIR_SNP_MIN: Minimum number of FIR SNPs. Default value = 75

MM_DIST: Number of Kbs between mismatches to end segment. Default = 1000.

NO_CALL: Character assigned to a no-call IN PHASED FILES.

© 2026 Mick Jolley (mickj1948@gmail.com)

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

# Path to .vcf file. 
VCF_FILE_PATH = r'********'

# Path to DNA files.
DNA_FILES_PATH = r'********' 

# Path to .xlsx file.
WORKING_DIRECTORY = r'********'

# Path to min_map.txt file.
MAP_PATH = r'*******'

# List of individuals to load from the VCF file. 
# Enter ['*'] to load all individuals. Leave empty ([]) to load none.
INDIVIDUALS = ['****','****','****']

# SUBJECTS to be compared from DNA_FILES_PATH. 
# Enter ['*'] to load all subjects. Leave empty ([]) to load none.
SUBJECTS = ['****','****','****']

# Chromosome selected. Leave empty to select all the chromosomes.
CHROMOSOMES = []

# Excel file name. Leave ".xlsx" out.
EXCEL_FILE_NAME = '*****'

# Percentage n0-calls allowed in files.
PC_NO_CALLS_ALLOWED = 10

# Calculate total cM in ROHs.
PARENTAL_RELATIONSHIP = False

# Suppress no-matches. Set to True if display of no-matches is desired.
SHOW_NO_MATCHES = True

# Chromosome true size. Set to False for normalized size.
CHROM_TRUE_SIZE = False

# Linearize the chromosome.
LINEAR_CHROMOSOME = False

# Resolution. Default = 1. Keep under 10. Set to 100 if full resolution is
# desired. If LINEAR_CHROMOSOME is set to True, RESOLUTION will be automatically
# set to 1, unless it is set to 10 (10x resolution).
RESOLUTION = 1

# HIR Minimum segment length (cM). The default is 7.
HIR_CUTOFF = 7

# FIR cutoff. FIRs less than 1cM in length are probably not significant.
FIR_CUTOFF = 1

# ROH cutoff (Mb).
ROH_CUTOFF = 5

# Display Fir tables.
FIR_TABLES = True

# Column to freeze. Set to "A" if freezing not required.
FREEZE_COLUMN = 'A'

# Linux font string. An alternative is:
# "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
LINUX_FONT_STRING = '*/fonts/truetype/family/DejaVuSerif-Bold.ttf'

"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""
You shouldn't have to change the parameters below.
"""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""""

# Minimum number of HIR SNPs.
HIR_SNP_MIN = 200

# Minimum number of FIR SNPs.
FIR_SNP_MIN = 75

# Number of Kbs between mismatches to end segment.
MM_DIST = 1000





