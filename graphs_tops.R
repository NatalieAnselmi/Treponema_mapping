#############################################
#Run First
#############################################
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
  "purrr",
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

#############################################
#Run second: Title by species
#############################################
var_title <- readline(prompt = "What is the species?: "); 

#############################################
#Run Graphing code blocks as needed
###################
#Upset Plots
###################
######### Top 50 Upset Plot #########
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

########## Top 100 Upset Plot #########
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
# COG Pie and stacked Bar charts
#############################################
# ---- Hardcoded palette ----
cog_palette <- c(
  "C"="#1f77b4","D"="#ff7f0e","E"="#2ca02c","F"="#d62728","G"="#9467bd",
  "H"="#8c564b","I"="#e377c2","J"="#7f7f7f","K"="#bcbd22","L"="#17becf",
  "M"="#aec7e8","N"="#ffbb78","O"="#98df8a","P"="#ff9896","Q"="#c5b0d5",
  "R"="#c49c94","S"="#f7b6d2","T"="#dbdb8d","U"="#9edae5","V"="#393b79",
  "No assignment"="grey80"
)

# ---- Data extraction ----
raw <- read_excel(excel_path, sheet = "Summary", col_names = FALSE)
raw_mat <- as.matrix(raw)
storage.mode(raw_mat) <- "character"

titles <- c(
  "Top50s","Top100s",
  "Top50vHealth_Higher","Top50vHealth_Lower",
  "Top100vHealth_Higher","Top100vHealth_Lower"
)

logical_matrix <- matrix(trimws(raw_mat) %in% titles, nrow = nrow(raw_mat))
title_positions <- which(logical_matrix, arr.ind = TRUE)

blocks <- list()
for (i in seq_len(nrow(title_positions))) {
  t_row <- title_positions[i,1]; t_col <- title_positions[i,2]
  title <- raw_mat[t_row, t_col]
  
  start_row <- t_row + 1
  start_col <- t_col
  
  # End row
  end_row <- start_row
  while (end_row <= nrow(raw_mat) &&
         !is.na(raw_mat[end_row,start_col]) &&
         raw_mat[end_row,start_col] != "") {
    end_row <- end_row + 1
  }
  end_row <- end_row - 1
  
  # End col
  end_col <- start_col + 1
  while (end_col <= ncol(raw_mat) &&
         !all(is.na(raw_mat[start_row:end_row,end_col])) &&
         !all(raw_mat[start_row:end_row,end_col] == "")) {
    end_col <- end_col + 1
  }
  end_col <- end_col - 1
  
  block <- raw_mat[start_row:end_row, start_col:end_col, drop = FALSE]
  colnames(block) <- block[1,]
  block <- block[-1, , drop=FALSE]
  block_df <- as.data.frame(block, stringsAsFactors = FALSE)
  
  cog_cat <- block_df[[1]]
  block_df <- block_df[,-1, drop=FALSE]
  block_df[] <- lapply(block_df, function(x) as.numeric(as.character(x)))
  block_df$COG_category <- cog_cat
  blocks[[title]] <- block_df
}

# ---- Plotting functions ----
# Pie chart builder
make_pie <- function(df, colname, title, cog_palette) {
  plot_df <- df %>%
    select(COG_category, !!sym(colname)) %>%
    rename(Count = !!sym(colname)) %>%
    filter(Count > 0) %>%
    arrange(desc(COG_category)) %>%
    mutate(
      Fraction = Count / sum(Count),
      ypos = cumsum(Fraction) - 0.5 * Fraction
    )
  
  ggplot(plot_df, aes(x = "", y = Fraction, fill = COG_category)) +
    geom_col(width = 1, color = "white") +
    coord_polar(theta = "y") +
    scale_fill_manual(values = cog_palette) +
    geom_text(
      aes(y = ypos, label = ifelse(Fraction > 0.02, Count, "")),
      color = "black", size = 3
    ) +
    labs(title = title) +
    theme_void() +
    theme(
      legend.position = "none",
      plot.title = element_text(hjust = 0.5, size = 12, face = "bold")
    )
}

# Shared legend for pies
get_shared_legend <- function(df, cols, cog_palette) {
  legend_df <- df %>%
    select(COG_category, all_of(cols)) %>%
    pivot_longer(cols = all_of(cols), names_to = "Group", values_to = "Value") %>%
    filter(Count > 0)
  
  legend_levels <- unique(legend_df$COG_category)
  
  p <- ggplot(legend_df, aes(x = "", y = Count, fill = COG_category)) +
    geom_col() +
    scale_fill_manual(values = cog_palette, breaks = legend_levels) +
    guides(fill = guide_legend(ncol = 2)) +
    theme_void() +
    theme(
      legend.position = "right",
      legend.key.size = unit(1.2, "lines"),
      legend.text = element_text(size = 10),
      legend.title = element_blank()
    )
  
  cowplot::get_legend(p) %>% patchwork::wrap_elements()
}

# Main plotting wrapper
plot_block <- function(block_name, df, plot_type = c("pie","bar"), palette = cog_palette) {
  plot_type <- match.arg(plot_type)
  
  # ---- Preserve original order ----
  # COG categories (rows of the block)
  df$COG_category <- factor(df$COG_category, levels = unique(df$COG_category))
  
  # Convert to long form
  df_long <- df %>%
    tidyr::pivot_longer(-COG_category, names_to = "Group", values_to = "Value")
  
  # Preserve original group order (columns of the block)
  df_long$Group <- factor(df_long$Group, levels = unique(df_long$Group))
  
  # ---- Pie charts ----
  if (plot_type == "pie") {
    pie_cols <- setdiff(names(df), "COG_category")
    pies <- lapply(pie_cols, function(col) {
      make_pie(df, colname = col, title = paste(block_name, "-", col), cog_palette)
    })
    shared_legend <- get_shared_legend(df, pie_cols, cog_palette)
    pie_grid <- wrap_plots(pies, ncol = 3) + 
      plot_layout(guides = "collect") & theme(legend.position = "none")
    return(pie_grid / shared_legend + plot_layout(heights = c(4, 1)))
  }
  
  # ---- Stacked bar ----
  if (plot_type == "bar") {
    return(
      ggplot(df_long, aes(x = Group, y = Value, fill = COG_category)) +
        geom_bar(stat = "identity") +
        scale_fill_manual(values = cog_palette) +
        ggtitle(bquote(italic(.(var_title)) ~ .(block_name) ~ "Stacked Bar")) +
        theme_minimal() +
        theme(
          axis.text.x = element_text(angle = 45, hjust = 1),
          legend.position = "right" )
    ) } }

#############################################
# ---- Choose which blocks to plot by commenting/uncommenting ----
plot_type <- "bar"   # <-- change to "pie" or "bar"

 plots1 <- plot_block("Top50s", blocks[["Top50s"]], plot_type = plot_type)
 plots2 <- plot_block("Top100s", blocks[["Top100s"]], plot_type = plot_type)
 plots3 <- plot_block("Top50vHealth_Higher", blocks[["Top50vHealth_Higher"]], plot_type = plot_type)
 plots4 <- plot_block("Top50vHealth_Lower", blocks[["Top50vHealth_Lower"]], plot_type = plot_type)
 plots5 <- plot_block("Top100vHealth_Higher", blocks[["Top100vHealth_Higher"]], plot_type = plot_type)
 plots6 <- plot_block("Top100vHealth_Lower", blocks[["Top100vHealth_Lower"]], plot_type = plot_type)

print(plots1)
print(plots2)
print(plots3)
print(plots4)
print(plots5)
print(plots6)
