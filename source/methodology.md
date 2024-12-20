## Methodology

The methodology for this project follows a structured approach to automate the generation of correlation and process capability (Cpk) reports:

**Step 1: Data Processing**
- Load the units datalog and limit board file from CSV.
- Perform data cleaning, format validation, and preprocessing to ensure data consistency.

**Step 2: Statistical Calculations**
- Compute key metrics for the units’ data, including:
  - Mean
  - Standard deviation
  - Delta mean
  - Mean shift
  - Standard deviation (SD) ratio
- Identify outliers or anomalies based on predefined thresholds for mean shift and SD ratio.

**Step 3: Cp and Cpk Calculations**
- Use the limit board file to calculate process capability metrics (Cp and Cpk) for each board.
- Compare the calculated values against industry benchmarks to evaluate quality control.

**Step 4: Report Generation**
- Automate the creation of a comprehensive report that includes:
  - Correlation results
  - Cp and Cpk values
- Format the output in a user-friendly layout, exportable to excel for easy distribution.
