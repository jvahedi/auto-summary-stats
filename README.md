# Auto Summary Statistics Generator
Author: John Vahedi  
Date: January 29, 2025

## Overview
This script is designed to automate the generation of summary statistics in accordance with the RAND style guide using a template Word document (`.docx`). By leveraging the `python-docx` library and a Pandas DataFrame, the script creates a formatted summary table that includes both categorical and numerical data statistics.

## Features
- **Customizable Parameters**: Allows users to adjust which features are treated as categorical versus numerical.
- **Automatic Table Population**: Automates the insertion and formatting of table headers, category labels, subcategories, and statistics.
- **RAND-Style Compliance**: Ensures that the generated document adheres to the formatting guidelines stipulated by the RAND organization.

## Usage
1. **Prepare Data**: Ensure your data is organized within a Pandas DataFrame (referred to in the script as `df`) that the script can access. This DataFrame should include columns intended for statistical summaries.

2. **Adjust Parameters**: Modify the `featured_categorical` dictionary to specify which columns should be treated as categorical (`True`) or numerical (`False`). Update the `headers` list if you wish to customize table column headers.

3. **Run the Script**: Execute the script in a Python environment with the necessary dependencies installed. This will process the data and overlay it onto the template Word document (`RAND_template.docx`), outputting a new file (`RAND_summary_statistics.docx`) with the formatted summary table.

## Configuration

### Parameter Example
Here's how you can configure your features and headers:

```python
# Significant digits for numerical statistics rounding
sig_dig = 1

# Feature Configuration: Specify whether each feature is categorical or numerical
featured_categorical = {
    'gender': True,            # Categorical Feature
    'main_discipline': True,   # Categorical Feature
    'journal_counts': False,   # Numerical Feature
    'publication_counts': False # Numerical Feature
    # Add more features as needed
}

# Table Headers Configuration: Customize the headers for your summary table
headers = ['Features', 'N (%)', 'Mean, Median, (Std. Dev.)\n[Range]']