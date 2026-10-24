# Netflix Dataset Cleaning 

## 1. Project Overview
The objective of this project was to clean and prepare the Netflix Titles dataset using Microsoft Excel before proceeding to the data visualization stage. The dataset contains information about Netflix movies and TV shows, including their titles, directors, cast members, countries, dates added to Netflix, release years, ratings, durations, genres, and descriptions.

The primary goal was to improve data quality, maintain consistency, handle missing values, check for duplicate records, and ensure that the dataset was ready for submission and approval before further analysis.

## 2. Tools Used
- **Microsoft Excel:** Data cleaning, filtering, missing-value handling, date conversion, and data validation.
- **Formulas:** `COUNTBLANK`, `ISNUMBER`, `IFERROR`, `DATE`, `MATCH`, and other Excel functions to check and transform data.

## 3. Data Cleaning Process

### A. Initial Data Inspection
The dataset was inspected to understand its structure, column names, data types, and overall quality. The main columns included `show_id`, `type`, `title`, `director`, `cast`, `country`, `date_added`, `release_year`, `rating`, `duration`, `listed_in`, and `description`.

This initial inspection helped identify missing values, inconsistent date formats, and areas requiring further validation.

### B. Identification and Treatment of Missing Values
Excel formulas and filters were used to identify missing values in important columns. The `COUNTBLANK` function helped count empty cells in the Director, Cast, Country, Date Added, Rating, and Duration columns.

Missing values in these columns were replaced with **"Unknown"** where appropriate. Excel's **Go To Special → Blanks** feature was used to select empty cells and fill them simultaneously using `Ctrl + Enter`. This process was applied to the relevant columns to reduce the number of blank cells and make missing information easier to identify.

### C. Date Format Standardization
The `date_added` column originally contained dates stored as text, such as *September 25, 2021*. These text values needed to be converted into actual Excel date values for consistent storage and future date-based analysis.

A temporary helper column was created to convert the text dates into valid Excel dates using formulas involving `DATE`, `MATCH`, `VALUE`, `LEFT`, `MID`, and `RIGHT`. The converted dates were formatted using Excel's Short Date format and then copied back into the original `date_added` column using **Paste Special → Values**. Missing dates were retained as "Unknown".

### D. Duplicate Record Check
The dataset was checked for duplicate records using Excel's duplicate-checking method. No duplicate records were found during the check. This helped confirm that the records did not contain obvious repeated entries according to the method used.

### E. Validation of the Type Column
The `type` column was reviewed using Excel filters. The expected categories, **Movie** and **TV Show**, were confirmed. These values distinguish the two main types of content available in the dataset and are useful for future comparisons.

### F. Validation of the Rating Column
The `rating` column was inspected to review the content classification values and identify missing or unexpected entries. Missing ratings were represented by "Unknown". Existing ratings were retained rather than unnecessarily changing valid classification categories.

### G. Validation of the Duration Column
The `duration` column was checked to identify missing values. Three blank duration cells were found and filled with "Unknown". The column contains different formats for movies and TV shows, such as minutes for movies and seasons for TV shows. These formats were retained because they represent different types of duration information.

### H. Verification of Missing Values
The `COUNTBLANK` function was used again to verify the relevant columns after cleaning. The goal was to confirm that no blank cells remained in the selected columns. A separate **Missing Values Check** worksheet was used to record the results and support the verification process.

### I. Temporary Column Management
Temporary helper columns used during duplicate checking and date conversion were reviewed after the cleaning operations. The intention was to retain the original dataset structure and remove unnecessary helper columns once the results had been verified.

## 4. Data Quality Improvements
The cleaning process focused on the following improvements:

- Reduced missing information in the selected columns by replacing blank cells with "Unknown".
- Standardized the `date_added` column by converting text dates into valid Excel dates.
- Checked for duplicate records.
- Reviewed the `type`, `rating`, and `duration` columns for consistency.
- Used formulas and filters to verify data quality.
- Maintained a separate worksheet for missing-value checks.

## 5. Outcome
The Netflix dataset was prepared for review and submission as a data-cleaning project. The work focused on improving consistency, handling missing information, and validating important columns while preserving the original records and useful information.

The cleaned dataset is intended to be submitted for approval before proceeding to the next stage of the project, which will involve data analysis and visualization.

## 6. Conclusion
This project provided practical experience in data cleaning using Microsoft Excel. By identifying missing values, converting text dates, checking duplicates, reviewing categorical fields, and verifying blank-cell counts, the dataset was made more consistent and easier to work with.

The cleaning stage establishes a foundation for reliable analysis. After the cleaned dataset is reviewed and approved, it can be used for data visualization to explore Netflix content trends, movie and TV show distributions, release-year patterns, ratings, countries, and other insights.
