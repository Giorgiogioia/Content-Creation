Hybrid Bulbs Template Processor - Description

Purpose:
This Python script generates a curated Excel file of products by merging template content with product data from a PIM (Product Information Management) system. It automatically fills placeholders in templates, evaluates conditional bullets, and applies consistent formatting to the output Excel file.

Key Steps and Functionality:

1. Paths Setup:

   * template\_path: Excel template containing placeholder text for product attributes.
   * pim\_path: Excel file containing product data including SKU and other details.
   * output\_path: Excel file listing SKUs to process.
   * output\_copy\_path: Path to save the final generated output file.

2. Helper Functions:

   * clean\_str(x): Normalizes text, removes special characters, and trims whitespace.
   * lower\_clean(x): Converts text to lowercase and cleans for consistent matching.
   * find\_ci(colnames, target): Performs case-insensitive search for column names.
   * detect\_max\_bullets\_ci(\*dfs): Determines the maximum number of bullet points across template and output files.

3. Data Loading and Cleaning:

   * Loads template, PIM, and output files as Pandas DataFrames.
   * Normalizes column names and key fields like SKU and Category.
   * Filters output rows to only include SKUs that exist in the PIM.

4. Placeholder Extraction and Processing:

   * extract\_placeholders(text): Identifies {placeholder} patterns within text.
   * process\_placeholder(placeholder, pim\_data, sku):

     * Supports simple replacements like {Column}.
     * Handles mapping placeholders in the form {"Column": {"key":"value", "default":""}}.
     * Supports join: placeholders to conditionally combine multiple fields.
   * process\_text(text, pim\_data, sku): Replaces all placeholders in standard text fields.

5. Bullet Point Processing:

   * process\_bullet(text, pim\_data, sku):

     * Evaluates bullet points with conditional switches: {switch: {"Column":"value"}} only keeps lines if PIM values match.
     * Provides fallback bullet if no condition matches.
     * Resolves any placeholders within bullet lines.

6. Column Resolution:

   * Maps template columns (Title, Bullet Points, Product Description) to existing or new output columns.
   * Adds missing columns to the output DataFrame as needed.

7. Main Processing Loop:

   * Iterates through each SKU in the output file.
   * Retrieves corresponding PIM data.
   * Selects the correct template row by Category.
   * Processes Title, Bullet Points, and Product Description using placeholders and conditional logic.
   * Writes the processed content back into the output DataFrame.

8. Saving to Excel with Formatting:

   * Uses xlsxwriter to save the final Excel file.
   * Applies column widths, text wrapping, alignment, and row heights.
   * Centers SKU and Category columns; sets Product Description to extra wide.
   * Formats headers with bold text and center alignment.

9. Logging and Warnings:

   * Displays SKUs missing in the PIM.
   * Logs placeholder processing and bullet condition evaluations.
   * Warns if no template row exists for a given category.

Output:

* A fully generated Excel file (Output\_Generated.xlsx) with all placeholders replaced, conditional bullets resolved, and consistent formatting.
* Original source files remain unchanged.
