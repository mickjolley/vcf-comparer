# -*- coding: utf-8 -*-
""" "
VCF_Comparer_configV1.py is the configuration file for VCF_Comparer.V1.0.py.

VCF_FILE_PATH: Add the .vcf file.
Example: VCF_FILE_PATH = r"c:/dna/xyx.vcf".

DNA_FILES_PATH: Path to folder where the DNA files are stored.

WORKING_DIRECTORY: Folder where the .xlsx and file will be stored.

MAP_PATH: Path to folder containing min_map.txt.

INDIVIDUALS: Add individuals from the .vcf file to be compared. Enter '*' if
all individuals are to be compared.

SUBJECTS: Rename subject file by inserting _raw_dna after name. Both .csv and
.txt files from all testing companies are accepted. PCV formatted files need not
be renamed.
Example: original file 37_S_Fred_123.csv. Renamedas 37_S_Fred_raw_dna123.csv.

CHROMOSOMES: Chromosome selected (1-23). More than one chromosome may be entered.
Enter '*' for all chromosomes.

EXCEL_FILE_NAME: Name of the .xlsx file. Do not include the ".xlsx", This is
added automatically.

SHOW_NO_MATCHES: Set to False if the display of match pairs with no matching
segments is not desired. This is the recommended default value. If all matches
are desired to be shown, set this to True.

CHROM_TRUE_SIZE: Set to True for true size. Set to False for normalized size.

LINEAR_CHROMOSOME: Set to True if you want to see the linearized chromosomes.
RESOLUTION will be ignored unless it is 10 (10x resolution). CHROM_TRUE_SIZE
must be set to False.

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

HIR_SNP_MIN: Minimum number of HIR SNPs. Default value = 200

FIR_SNP_MIN: Minimum number of FIR SNPs. Default value = 75

MM_DIST: Number of Kbs between mismatches to end segment. Default = 1000.

NO_CALL: Character to designate a no call.

© 2026 Mick Jolley (mickj1948@gmail.com)

"""

# Path to vcf file. Add .vcf file to the end
VCF_FILE_PATH = r"c:/dna/v62.HO.GBR.Orkney.DG.vcf"

# Path to SUBJECTS DNA files. Add .vcf file to the end if .vcf file is to be processed.
DNA_FILES_PATH = r"c:/dna files"

# Path to .xlsx file.
WORKING_DIRECTORY = r"c:/vpphaser"

# Path to min_map.txt file.
MAP_PATH = r"c:/minmap"

# Individuals from .vcf files to be compared.
INDIVIDUALS = ["HG00096.DG", "HG00097.DG", "HG00110.DG", "HG00111.DG", "HG00115.DG"]

# SIBLINGS to be compared. Make sure that no two files share the same name.
SUBJECTS = ["Jean", "Christine", "Mick"]

# Chromosome selected. Enter '*' to select all the chromosomes.
CHROMOSOMES = ["*"]

# Excel file name. Leave ".xlsx" out.
EXCEL_FILE_NAME = "test16"

# Suppress no-matches. Set to True if display of no-matches is desired.
SHOW_NO_MATCHES = False

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

# X chromosome HIR Minimum segment length (cM). The default is 15.
X_HIR_CUTOFF = 15

# X chromosome FIR cutoff. FIRs less than 15cM in length are probably not significant.
X_FIR_CUTOFF = 15

# Turn scale on and off. Set to False if not required.
SCALE_ON = True

# Column to freeze. Set to "A" if freezing not required.
FREEZE_COLUMN = "A"

# Linux font string. An alternative is:
# "/usr/share/fonts/truetype/dejavu/DejaVuSerif-Bold.ttf"
LINUX_FONT_STRING = "*/fonts/truetype/family/DejaVuSerif-Bold.ttf"


"""""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""
You shouldn't have to change the parameters below.
""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""" """""" ""

# Minimum number of HIR SNPs.
HIR_SNP_MIN = 200

# Minimum number of FIR SNPs.
FIR_SNP_MIN = 75

# Number of Kbs between mismatches to end segment.
MM_DIST = 1000

# Character assigned to no calls in phased files.
NO_CALL = "0"
