# ===== Install & Load Packages =====
required_packages <- c(
  "readxl",      # read Excel files
  "tidyverse",   # data wrangling + ggplot2
  "pheatmap",    # heatmaps
  "FactoMineR",  # PCA
  "factoextra",  # PCA visualizations
  "scales"       # non scientific axis
)

# Install missing CRAN packages
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg)
  }
  library(pkg, character.only = TRUE) # Load package after install
}

# Install EnhancedVolcano from Bioconductor if not already installed
if (!requireNamespace("EnhancedVolcano", quietly = TRUE)) {
  if (!requireNamespace("BiocManager", quietly = TRUE)) {
    install.packages("BiocManager")
  }
  BiocManager::install("EnhancedVolcano")
}
library(EnhancedVolcano)

# ===== Prompt user to select Excel workbook =====
file_path <- file.choose()

# Get sheet names and filter to only health status sheets
all_sheets <- excel_sheets(file_path)
health_sheets <- setdiff(all_sheets, c("Full_Data", "Unmatched", "Summary", "Pos. Genes", "Top50s"))

# ===== Read all health status sheets into a single long-format dataframe =====
all_data <- map_dfr(health_sheets, function(sheet) {
  df <- read_excel(file_path, sheet = sheet)
  
  # Ensure Mapped_Health_Status column exists; if not, create from sheet name
  if (!"Mapped_Health_Status" %in% colnames(df)) {
    df <- df %>% mutate(Mapped_Health_Status = sheet)
  }
  
  # Pivot to long format
  df_long <- df %>%
    pivot_longer(
      cols = -c(Sample, Mapped_Health_Status),
      names_to = "Gene",
      values_to = "Abundance"
    )
  return(df_long)
})

# Ensure factor levels 
all_data$Mapped_Health_Status <- factor(all_data$Mapped_Health_Status, levels = health_sheets)

#####################################
# 3. Heatmap: Mean gene abundance per health status top 50 - CHECKED
#####################################
# Top 50 variable genes
top_genes_50 <- all_data %>%
  group_by(Gene) %>%
  summarise(var = var(Abundance, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(var)) %>%
  slice_head(n = 50) %>%
  pull(Gene)

# Filter for top 50 genes
wide_top50 <- all_data %>% filter(Gene %in% top_genes_50)

# Compute mean abundance per gene per health status
mean_abundance <- wide_top50 %>% 
  group_by(Gene, Mapped_Health_Status) %>% 
  summarise(MeanAbundance = mean(Abundance, na.rm = TRUE), .groups = "drop") %>% 
  pivot_wider(names_from = Mapped_Health_Status, values_from = MeanAbundance) 

# Set explicit column order for health statuses
health_order <- c("Healthy", "PD", "Stable PD", "Fluctuating PD", "Progressing PD")
mean_abundance <- mean_abundance[, c("Gene", health_order)]

# Convert Gene to rownames
mean_abundance <- column_to_rownames(mean_abundance, var = "Gene")

# Plot heatmap
p1 <- pheatmap(
  log1p(as.matrix(mean_abundance)),
  cluster_rows = TRUE,      # genes are clustered by similarity of abundance across health statuses
  cluster_cols = FALSE,     # columns are kept in the specified order
  main = "Heatmap of mean gene abundance: 50 most variable genes (log1p)"
)
print(p1)

#####################################
# 3. Heatmap: Mean gene abundance per health status top 100 - CHECKED
#####################################
# Top 100 variable genes
top_genes_100 <- all_data %>%
  group_by(Gene) %>%
  summarise(var = var(Abundance, na.rm = TRUE), .groups = "drop") %>%
  arrange(desc(var)) %>%
  slice_head(n = 100) %>%
  pull(Gene)

# Filter for top 100 genes
wide_top100 <- all_data %>% filter(Gene %in% top_genes_100)

# Compute mean abundance per gene per health status
mean_abundance <- wide_top100 %>% 
  group_by(Gene, Mapped_Health_Status) %>% 
  summarise(MeanAbundance = mean(Abundance, na.rm = TRUE), .groups = "drop") %>% 
  pivot_wider(names_from = Mapped_Health_Status, values_from = MeanAbundance) 

# Set explicit column order for health statuses
health_order <- c("Healthy", "PD", "Stable PD", "Fluctuating PD", "Progressing PD")
mean_abundance <- mean_abundance[, c("Gene", health_order)]

# Convert Gene to rownames
mean_abundance <- column_to_rownames(mean_abundance, var = "Gene")

# Plot heatmap
p2 <- pheatmap(
  log1p(as.matrix(mean_abundance)),
  cluster_rows = TRUE,      # genes are clustered by similarity of abundance across health statuses
  cluster_cols = FALSE,     # columns are kept in the specified order
  main = "100 most variable gene abundance (mean; log1p)"
)
print(p2) #save 460w x 1200h

#####################################
# PCA of samples using all genes, genes colored by health status they contribute most to - Okay
#####################################
# Aggregate the data by gene and health status
gene_summary <- all_data %>%
  group_by(Gene, Mapped_Health_Status) %>%
  summarise(mean_abundance = mean(Abundance, na.rm = TRUE), .groups = "drop") %>%
  pivot_wider(names_from = Mapped_Health_Status, values_from = mean_abundance, values_fill = 0)

# Prepare matrix for PCA
# Rows = genes, columns = health statuses
gene_matrix <- gene_summary %>%
  column_to_rownames("Gene") %>%
  as.matrix()

# Determine dominant health status for each gene
gene_dominant_status <- apply(gene_matrix, 1, function(x) colnames(gene_matrix)[which.max(x)])

# Define colors
status_levels <- c("Healthy", "PD", "Stable PD", "Fluctuating PD", "Progressing PD")
gene_dominant_status <- factor(gene_dominant_status, levels = status_levels)
# Assign colors for points and pale fills for ellipses
point_colors <- RColorBrewer::brewer.pal(length(status_levels), "Set1")
ellipse_fills <- scales::alpha(point_colors, 0.2)  # pale for ellipses

# Run PCA
res_pca <- PCA(gene_matrix, graph = FALSE)

# Determine dominant health status for each gene
gene_dominant_status <- apply(gene_matrix, 1, function(x) colnames(gene_matrix)[which.max(x)])

# Ensure gene_dominant_status is a factor with desired order
gene_dominant_status <- factor(gene_dominant_status, 
                               levels = c("Healthy", "PD", "Stable PD", "Fluctuating PD", "Progressing PD"))

# Extract percentage of variance explained from PCA
var_explained <- res_pca$eig[, 2]  # column 2 contains % variance explained
pc1_label <- paste0("PC1 – ", round(var_explained[1], 1), "% variance explained")
pc2_label <- paste0("PC2 – ", round(var_explained[2], 1), "% variance explained")

# Plot
p3 <- fviz_pca_ind(res_pca,
                        geom.ind = "point",
                        col.ind = gene_dominant_status,   # color points by dominant health status
                        palette = "jco",
                        addEllipses = TRUE,
                        repel = FALSE) +
  theme_minimal(base_size = 14) +
  labs(color = "Dominant Health Status",
       title = "Gene-level PCA by Health Status",
       x = pc1_label,
       y = pc2_label) +
  guides(shape = "none") +               # remove extra shape legend
  guides(fill = "none")
print(p3)

#####################################
# Four Volcano plots (each disease vs Healthy) - CHECKED
#####################################
disease_sheets <- health_sheets[health_sheets != "Healthy"] 
volcano_plots <- list()

for(disease in disease_sheets){ 
  gene_stats <- all_data %>% 
    filter(Mapped_Health_Status %in% c("Healthy", disease)) %>% 
    group_by(Gene) %>% 
    summarise( 
      log2FC = log2(mean(Abundance[Mapped_Health_Status == disease] + 1e-6) /
                      mean(Abundance[Mapped_Health_Status == "Healthy"] + 1e-6)), 
      pval = t.test(Abundance ~ Mapped_Health_Status)$p.value )
  
  volcano_plots[[disease]] <- EnhancedVolcano( 
    gene_stats, 
    lab = gene_stats$Gene, 
    x = "log2FC", 
    y = "pval", 
    pCutoff = 0.05, 
    FCcutoff = 1, 
    pointSize = 2, 
    labSize = 3, 
    title = paste(disease, "vs Healthy") 
  )
}
for(p in volcano_plots) 
  print(p) 

#####################################
# TOP 20 VARIABLE GENES - Boxplot - CHECKED
#####################################
#top_genes <- all_data %>%
#  group_by(Gene) %>%
#  summarise(var = var(Abundance, na.rm = TRUE), .groups = "drop") %>%
#  arrange(desc(var)) %>%
#  slice_head(n = 20) %>%
#  pull(Gene)

#wide_top20 <- all_data %>% filter(Gene %in% top_genes)

#p4 <- ggplot(wide_top20, aes(x = Gene, y = Abundance + 1e-6, fill = Mapped_Health_Status)) +
#  geom_boxplot(outlier.shape = NA, width = 0.7) +
#  scale_y_log10() +
#  theme_minimal() +
#  theme(axis.text.x = element_text(angle = 45, hjust = 1)) +
#  labs(title = "Abundance of Top 20 Variable Genes", y = "Abundance (log10 scale)")

#print(p4)

#####################################
# Density plot: Abundance distribution - CHECKED
# NOT NEEDED
#####################################
#p2 <- ggplot(all_data, aes(x = Abundance + 1e-6, fill = Mapped_Health_Status)) + 
#  geom_density(alpha = 0.4) + 
#  scale_x_log10() + 
#  labs(x = "Gene Abundance (log scale)", y = "Density") + 
#  theme_bw()
#print(p2)