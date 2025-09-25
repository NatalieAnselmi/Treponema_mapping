# Treponema_mapping

**Set up**  
	1.	If data in text file, convert to excel file  
	  
	2.	Go to https://www.ncbi.nlm.nih.gov/home/genomes/ and search for species and strain of interest  
	    ⁃	Download —> Genebank only —> Sequence and annotation (GBFF) only  
	    ⁃	unzip download, find the .gbff file (nested a few folders in) and move it to your work folder  
	  
	3.	Make sure you have a metadata folder listing the health status of each sample  
	  
	4.	Run genome_to_fasta.py to convert the .gbff into a cleanly FASTA formatted full genome text file. Can select whether you would like sequences named by gene locus or protein product (repeated named like ‘hypothetical protein’ will be assigned an integer)  
	    ⁃	Go to http://eggnog-mapper.embl.de/ and submit the full genome FASTA  
	    ⁃	Download as excel file and rename as needed  

**Data Analysis**
	1.	Run sort_and_sumarize.py  to create a new workbook (Mapped Data) with data in sheets based on health status  
	    ⁃	Requires  
	       ⁃	Excel data table with sample names, gene names/loci, reads  
	       ⁃	Sample Metadata table with sample names and health status  
	       ⁃	If converting from gene locus to gene product, requires .gbff file  
	  
	2.	Open and check on file to ensure code ran correctly  
	  
	3.	Get important gene populations from Mapped Data workbook  
	    ⁃	top_genes_COGs.py will create a new workbook (TopGenes)  
	    ⁃	1 sheet of top 50 most abundant genes per health status + top 50 most variable genes across health statuses  
	    ⁃	1 sheet of top 100 most abundant genes per health status + top 100 most variable genes across health statuses  
	    ⁃	1 sheet of top 50 or 100 (log₂ fold-change vs Healthy with pseudocount = 0.1; computes both higher and lower for 8 columns a sheet)  
	    ⁃	Summary sheet with tables of COGs per condition  
  
**Data Visualization**
	1.	graphs_overall.R has code to produce  
	⁃	Mean gene abundance per health status (top 50 and top 100)  
	⁃	PCA of all samples (somewhat useful)  
	⁃	Volcano plots for each disease state vs Healthy  
	⁃	Commented out code for box plot of top 20 variable genes and density plot of all data  
	  
	2.	graphs_tops.R has code to produce:  
	⁃	Upset plots of top 100 and top 50 gene lists  
	⁃	COG pie charts (1 per column of a given summary table)  
	⁃	One stacked bar per summary table  
