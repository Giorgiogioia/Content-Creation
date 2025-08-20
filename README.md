Hybrid Bulbs Template Processor

This script automates the generation of a curated product Excel file by combining template content with product data from a PIM (Product Information Management) system. It fills placeholders in templates, processes conditional bullets, and applies formatting to the output Excel file.

Key Steps and Functionality:

Paths Setup:

template_path: Excel template containing text and placeholders for product content.

pim_path: Excel file containing product data (SKU, attributes, etc.).

output_path: Curated list of SKUs to process.

output_copy_path: Path to save the generated output Excel file.

Helper Functions:

clean_str(x): Normalizes Unicode, removes special spaces, and trims text.

lower_clean(x): Lowercases and cleans text for case-insensitive matching.

find_ci(colnames, target): Finds column names case-insensitively.

detect_max_bullets_ci(*dfs): Determines the maximum number of bullet points in the template or output.

Data Loading and Cleaning:

Loads template, PIM, and output files as Pandas DataFrames.

Normalizes column names and required keys (SKU, Category).

Filters the output DataFrame to only include SKUs found in the PIM.

Placeholder Extraction and Processing:

extract_placeholders(text): Extracts {placeholder} patterns from text.

process_placeholder(placeholder, pim_data, sku):

Handles simple replacements {Column}.

Supports mapping placeholders: {"Column": {"key": "value", "default": ""}}.

Supports join: placeholders to combine multiple fields conditionally.

process_text(text, pim_data, sku): Replaces all placeholders in plain text.

Bullet Point Processing:

process_bullet(text, pim_data, sku): Handles bullet lines with conditional switches:

{switch: {"Column": "value"}} keeps the line only if the PIM value matches.

Supports fallback bullets if no switch matches.

Resolves placeholders within bullets.

Column Resolution:

Matches template columns (Title, Bullet Points, Product Description) to output columns.

Adds missing columns in the output DataFrame if necessary.

Main Processing Loop:

Iterates through each SKU in the output file.

Retrieves PIM data for the SKU.

Finds the matching template row by Category.

Processes Title, Bullet Points, and Product Description using placeholders and conditional logic.

Writes processed content back to the output DataFrame.

Saving to Excel with Formatting:

Uses xlsxwriter to save the output file.

Applies column widths, text alignment, text wrapping, and row heights.

Centers SKU and Category columns; makes Product Description extra wide.

Formats header row with bold, centered text.

Logging and Warnings:

Prints info about SKUs missing from the PIM.

Shows which placeholders, mappings, or switch bullets are processed.

Warns if template rows are missing for a category.

Output:

A fully generated Excel file (Output_Generated.xlsx) with all placeholders replaced, conditional bullets resolved, and formatting applied.

Source files (template, PIM, output) remain unchanged.
