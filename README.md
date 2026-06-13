
# VCF_Comparer.V3.0.py.

## Overview
VCF Comparer is a high-performance bioinformatics tool designed to identify **Half-Identical Regions (HIR)** 
and **Fully-Identical Regions (FIR)** between individuals. It supports input from multi-sample VCF files and 
individual raw DNA files (e.g. Ancestry, 23andMe).

VCF_Comparer.V3.0.py has a number of advanced features.

1) PC_NO_CALLS_ALLOWED. This option gives the user the ability set a threshold on the percentage of no-calls in a .vcf 
   file. The default value (for moderns) is 10%. This value may be increased for ancient samples (try 50% to start). 
   This will filter out files of poor quality. The default value is 10%.

2) PARENTAL_RELATIONSHIP. This option calculates the total cMs of runs of homozygosity. Multiply this value by four to 
   get the approximate equivalent of match values.

Sample .vcf files are provided. Orkney.vcf is a modern group of individuals (1000 Genomes). Endogamy is prevalent in this 
population. Ancients.vcf is composed of a mixture of ancient individuals. To find out more about them consult this website:
https://amtdb.org/samples.

---

## 1. Setup & Requirements

### Prerequisites
- **Python 3.8+**
- **Required Libraries:**
  ```bash
  pip install numpy pandas pillow openpyxl
  ```

### File Structure
- `VCF_Comparer.V3.0.py`: The main execution script.
- `VCF_Comparer_configV3.py`: Configuration file for paths and thresholds.
- `min_map.txt`: A tab-delimited genetic map file (required for cM calculations).

---

## 2. Configuration (`VCF_Comparer_configV3.py`)

Edit the configuration file to set your paths and analysis parameters:

### Paths
- `VCF_FILE_PATH`: Path to your `.vcf` file.
- `DNA_FILES_PATH`: Folder containing raw DNA files. Files must contain `_raw_dna` after the name 
                    (e.g., `XYZ_John_raw_dna.csv`). Both .csv and .txt files from all testing companies
					can be used.
- `WORKING_DIRECTORY`: Where results are saved.
- `MAP_PATH`: Folder containing `min_map.txt`.

### Individual Selection
-  'INDIVIDUALS = asterisk': Loads all samples from the VCF.
-  'INDIVIDUALS = ["SampleA", "SampleB"]`: Loads specific samples from the VCF file.
-  'SUBJECTS = asterisk': Loads all files from `DNA_FILES_PATH`.
-  'SUBJECTS = ["PersonX", "PersonY"]': Loads specific individual files.
- **Comparison Logic:**
  - If 'SUBJECTS' and 'INDIVIDUALS' are provided, they are compared against 'INDIVIDUALS'.
  - If 'SUBJECTS' is empty, 'INDIVIDUALS' are compared against each other.
  - If 'INDIVIDUALS' is empty, 'SUBJECTS' are compared against each other.
  
### Analysis Thresholds
- `HIR_CUTOFF`: Minimum cM length for HIR segments (e.g., `7`).
- `FIR_CUTOFF`: Minimum cM length for FIR segments (e.g., `1`).
- `MM_DIST`: The "stopper" distance in Kb. If two mismatches are closer than this, the segment ends. (Default: `1000`).
- `HIR_SNP_MIN` / `FIR_SNP_MIN`: Minimum SNPs required to validate a segment.

---

## 3. Visualization Modes

VCF_Comparer.V3.0.py offers three distinct ways to visualize chromosomes in the Excel report:

### A. SNP-Based (Default)
- **Setting:** `CHROM_TRUE_SIZE = False`, `LINEAR_CHROMOSOME = False`
- **Behavior:** The width of the chromosome image is normalized so that each chromosome is the same pixel length.

### B. Linear Base-Pair Scaling
- **Setting:** `CHROM_TRUE_SIZE = False`, `LINEAR_CHROMOSOME = True`
- **Behavior:** Chromosomes are scaled linearly by their physical base-pair length. Positions with no data are colored grey.

### C. True Size (Physical Proportions)
- **Setting:** `CHROM_TRUE_SIZE = True`
- **Behavior:** Chromosomes are drawn proportional to their actual size (e.g., Chromosome 1 is much wider than Chromosome 22).

In additition, 'SHOW_NO_MATCHES = 'False' suppresses the display of no matches.
  
---

## 4. Understanding the Output

### Excel Report (`.xlsx`)
- **Tabs:** Each chromosome has its own sheet (Chr1, Chr2, etc.).
- **Segment Tables:** Lists every matching segment found for every pair, including Start/Finish Mb, SNP count, and cM length.
- **Visual Plots:**
  - **Top Bar (Colors):**
    - Green: Full match (FIR).
    - Yellow: Half match (HIR).
    - Red: Nomatch (NIR).
    - Grey: No data for one or both individuals.
  - **Bottom Bar (Blocks):**
    - Blue: Validated HIR segment (above cutoff).
    - Orange: Validated FIR segment (above cutoff).
	- Black: NIR segment.

### .csv Summary
- A summary of total cM shared between all pairs across all analyzed chromosomes.

### _ROH.csv Summary
- A summary of total cM of ROHs in individuals' DNA files. Multiply by four to approximate matches shared between two individuals.

---

## 5. Performance Tips
- **VCF Size:** The script uses vectorized lookups for VCF parsing. For very large files, ensure you have enough 
                RAM (roughly 4x the size of the VCF is recommended).
- **Parallel Processing:** The script automatically utilizes all available CPU cores for chromosome analysis.
- **Resolution:** `RESOLUTION` can be set from 1 (approximately 1000 pixels) to 100 (full resolution). If RESOLUTION x 1000
                  is greater than the length of the chromosome, full resolution is the result.

---

## 6. Error Reporting
The tool validates your inputs before starting:
- **Missing Files:** It will explicitly name which VCF, Map, or Directory is missing.
- **Missing Individuals:** If a name in your config doesn't match a VCF header or a file on disk, a warning is printed immediately.
- **Empty Data:** If an individual has no data for a specific chromosome, they are skipped for that chromosome without crashing.

---
© 2026 Mick Jolley. All rights reserved.
