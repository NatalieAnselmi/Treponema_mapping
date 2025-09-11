# --- Helper: Install & load packages ---
packages <- c(
  "readxl", 
  "dplyr", 
  "ggplot2", 
  "tidyr", 
  "patchwork",
  "tidyverse",
  "UpSetR",
  "paletteer",
  "cowplot",
  "ggrepel")

install_if_missing <- function(pkg){
  if (!requireNamespace(pkg, quietly = TRUE)) {
    install.packages(pkg, repos = "https://cloud.r-project.org")
  }
  suppressPackageStartupMessages(library(pkg, character.only = TRUE))
}

invisible(lapply(packages, install_if_missing))

# ---- Select the Excel file interactively ----
cat("Select your Top listed genes excel file in the dialog box...\n")
excel_path <- file.choose()

#Title by species
var_title <- readline(prompt = "What is the species?: "); 

#############################################
#Upset Plots
##############################################
# Top 50 Upset Plot
# ---- Read the first sheet ----
df <- read_excel(excel_path, sheet = "Top50s")

# ---- Convert each column into a list of genes ----
gene_lists <- lapply(df, function(col) unique(na.omit(as.character(col))))

# ---- Convert lists into a presence/absence dataframe for UpSetR ----
all_genes <- unique(unlist(gene_lists))
upset_data <- data.frame(Gene = all_genes)

for (name in names(gene_lists)) {
  upset_data[[name]] <- ifelse(upset_data$Gene %in% gene_lists[[name]], 1, 0)
}

# ---- Make the UpSet plot ----
upset(
  upset_data,
  sets = names(gene_lists),
  nsets = length(gene_lists),
  order.by = "freq",
)

#############################
# Top 100 Upset Plot
# ---- Read the second sheet ----
df <- read_excel(excel_path, sheet = "Top100s")

# ---- Convert each column into a list of genes ----
gene_lists <- lapply(df, function(col) unique(na.omit(as.character(col))))

# ---- Convert lists into a presence/absence dataframe for UpSetR ----
all_genes <- unique(unlist(gene_lists))
upset_data <- data.frame(Gene = all_genes)

for (name in names(gene_lists)) {
  upset_data[[name]] <- ifelse(upset_data$Gene %in% gene_lists[[name]], 1, 0)
}

# ---- Make the UpSet plot ----
upset(
  upset_data,
  sets = names(gene_lists),
  nsets = length(gene_lists),
  order.by = "freq",
)

#############################################
# COG Pie charts (50 and 100)
#############################################
# ---- Read the third sheet ----
summary_df <- read_excel(excel_path, sheet = "Summary")

# Make sure COG_category is character
summary_df <- summary_df %>%
  dplyr::rename(COG = COG_category) %>%
  dplyr::mutate(COG = as.character(COG))

# Define consistent colors for all COGs
cog_codes <- unique(summary_df$COG)
palette <- setNames(
  paletteer::paletteer_d("ggthemes::Classic_20"),
  #paletteer::paletteer_d("ggthemes::manyeys"),  #alt color palette
  cog_codes
)

# --- Updated legend builder with formatting ---
get_shared_legend <- function(df, cols) {
  # Gather all counts from the selected columns
  legend_df <- df %>%
    dplyr::select(COG, all_of(cols)) %>%
    tidyr::pivot_longer(cols = all_of(cols), 
                        names_to = "Condition", 
                        values_to = "Count") %>%
    dplyr::filter(Count > 0)
  legend_levels <- unique(legend_df$COG)

  # Dummy plot to extract a legend
  p <- ggplot(legend_df, aes(x = "", y = Count, fill = COG)) +
    geom_col() +
    scale_fill_manual(
      values = palette, 
      breaks = legend_levels) +
    guides(
      fill = guide_legend(ncol = 2)  # two-column layout
    ) +
    theme_void() +
    theme(
      legend.position = "right",
      legend.key.size = unit(1.2, "lines"),
      legend.text = element_text(size = 10),
      legend.title = element_blank()
    )
  
  # Extract and wrap the legend so it can expand in patchwork
  leg <- cowplot::get_legend(p)
  return(patchwork::wrap_elements(leg))
}

# Helper functions to make chart
make_pie <- function(df, colname, title) {
  plot_df <- df %>%
    dplyr::select(COG, !!sym(colname)) %>%
    dplyr::rename(Count = !!sym(colname)) %>%
    dplyr::filter(Count > 0) %>%
    dplyr::arrange(desc(COG)) %>%
    dplyr::mutate(
      Fraction = Count / sum(Count),
      ypos = cumsum(Fraction) - 0.5 * Fraction
    )
  
  ggplot(plot_df, aes(x = "", y = Fraction, fill = COG)) +
    geom_col(width = 1, color = "white") +
    coord_polar(theta = "y") +
    scale_fill_manual(values = palette, breaks = cog_levels) +
    geom_text(
      aes(y = ypos, label = ifelse(Fraction > 0.02, Count, "")),  # show only if >2%
      color = "black", size = 3
    ) +
    labs(title = title) +
    theme_void() +
    theme(
      legend.position = "none",
      plot.title = element_text(hjust = 0.5, size = 12, face = "bold")
    )
}

#############################################
# --- Build Top 50 set ---
top50_cols <- c("Top50s_Healthy", "Top50s_PD", "Top50s_Stable PD",
                "Top50s_Fluctuating PD", "Top50s_Progressing PD")

top50_charts <- list(
  make_pie(summary_df, "Top50s_Healthy", "Healthy"),
  make_pie(summary_df, "Top50s_PD", "PD"),
  make_pie(summary_df, "Top50s_Stable PD", "Stable PD"),
  make_pie(summary_df, "Top50s_Fluctuating PD", "Fluctuating PD"),
  make_pie(summary_df, "Top50s_Progressing PD", "Progressing PD")
)
top50_legend <- get_shared_legend(summary_df, top50_cols)

top50_plot <- (top50_charts[[1]] | top50_charts[[2]] | top50_charts[[3]]) /
  (top50_charts[[4]] | top50_charts[[5]] | top50_legend) +
  plot_annotation(
    title = bquote(
      bold("Top 50 Most Abundant Gene COGs - ") ~ bolditalic(.(var_title))
    )) +
  theme(plot.title = element_text(size = 16, hjust = 0.5))

# --- Build Top 100 set ---
top100_cols <- c("Top100s_Healthy", "Top100s_PD", "Top100s_Stable PD",
                 "Top100s_Fluctuating PD", "Top100s_Progressing PD")

top100_charts <- list(
  make_pie(summary_df, "Top100s_Healthy", "Healthy"),
  make_pie(summary_df, "Top100s_PD", "PD"),
  make_pie(summary_df, "Top100s_Stable PD", "Stable PD"),
  make_pie(summary_df, "Top100s_Fluctuating PD", "Fluctuating PD"),
  make_pie(summary_df, "Top100s_Progressing PD", "Progressing PD")
)
top100_legend <- get_shared_legend(summary_df, top100_cols)

top100_plot <- (top100_charts[[1]] | top100_charts[[2]] | top100_charts[[3]]) /
  (top100_charts[[4]] | top100_charts[[5]] | top100_legend) +
  plot_annotation(
    title = bquote(
      bold("Top 100 Most Abundant Gene COGs - ") ~ bolditalic(.(var_title))
    )) +
  theme(plot.title = element_text(size = 16, hjust = 0.5))

######################
#Variable Sets
######################
# --- Build Var 50 set ---
var50_col <- c("Top50s_Most Variable")
var50_chart <- make_pie(summary_df, "Top50s_Most Variable", NULL)
var50_legend <- get_shared_legend(summary_df, var50_col)

var50_chart <- 
  (var50_chart | var50_legend) +
  plot_annotation(title = paste(var_title)) &
  theme(plot.title = element_text(size = 16, face = "bold.italic", hjust = 0.5))

# --- Build Var 100 set ---
var100_col <- c("Top100s_Most Variable")
var100_chart <- make_pie(summary_df, "Top100s_Most Variable", NULL)
var100_legend <- get_shared_legend(summary_df, var100_col)

var100_chart <- 
  (var100_chart | var100_legend) +
  plot_annotation(title = paste(var_title)) &
  theme(plot.title = element_text(size = 16, face = "bold.italic", hjust = 0.5))

# --- Show plots ---
print(top50_plot)
print(top100_plot)
print(var50_chart)
print(var100_chart)