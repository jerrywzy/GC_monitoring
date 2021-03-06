## Overview

Script to generate plots to monitor PCR replicates, NC, QC scores and Average Scores for GC

## Instructions 
1. Prepare directories - gc_monitoring.py will extract data from all \*/\*.xlsm files.

2. Run gc_tabulation.py 
	- Outputs CSV files for PCR replicates (G1-12), NC scores (G1-12), QR scores and Average Scores Per Sample
	
3. Run gc_plots.py
	- Outputs graphs for each CSV file above

3. The following graphs are generated in the GC_monitoring/ folder:
    - PCR replicate stripplot for each gene from G1-12, for each run
    - NC stripplot for each gene from G1-12, for each run
    - Scores for QR1 and QR2 for each run
    - Average Scores for samples for each run

## Required libraries
  - Pandas
  - Numpy
  - matplotlib
  - Seaborn
  - glob
  
