# Assignment-3
Creating a combined sheet
Medical Data Cleaning
Overview
This project focuses on cleaning and preparing medical and hospitalization data using Microsoft Excel. The dataset contains two main tables:

Medical Examinations
Hospitalization Details
The primary objective is to identify and handle missing values represented by ?, replace them with appropriate values, and improve the overall quality of the dataset before further analysis.

Data Cleaning Tasks
1. Identify Missing Values
The first step was to check the number of missing values represented by ? in each column of both the Medical Examinations and Hospitalization Details tables.

The COUNTIF function was used to count missing values.

=COUNTIF(range,"?")
Example:

=COUNTIF('Hospitalisation Details'!A:A,"?")
This formula counts all cells containing ? in the selected column.

A separate solution table was created to summarize the number of missing values found in each column.

2. Handle Missing Month and Year Values
Missing values in the month and year columns were handled separately.

Month
Missing values in the month column were replaced with:

Sep
This was done using Excel's Find and Replace feature:

Select the month column.
Press Ctrl + H.
In Find what, enter ?.
In Replace with, enter Sep.
Click Replace All.
Year
For the year column, missing values were replaced with the average year rounded to the nearest integer.

The average was calculated using:

=ROUND(AVERAGE(B:B),0)
After calculating the rounded average, the missing ? values in the year column were replaced with the calculated value.

3. Fill Missing Values Using the Most Frequent Value
Missing values in the following columns were filled using their most frequently occurring value (mode):

smoker
Hospital tier
City tier
Hospital Tier
The frequency of each hospital tier was calculated using COUNTIF.

=COUNTIF('Hospitalisation Details'!G:G,"tier - 1")
=COUNTIF('Hospitalisation Details'!G:G,"tier - 2")
=COUNTIF('Hospitalisation Details'!G:G,"tier - 3")
The tier with the highest count was identified as the most frequently occurring value and used to replace missing values.

City Tier
The same method was applied to the City tier column by calculating the frequency of each tier and selecting the one with the highest occurrence.

Smoker
The smoker column contains two main values:

yes
no
A Pivot Table was used to calculate their frequencies.

Steps:

Select the dataset.
Go to Insert → PivotTable.
Add smoker to the Rows field.
Add smoker again to the Values field and use Count.
Compare the counts of yes and no.
Replace missing values with the most frequently occurring value.
4. Handle Missing State ID Values
Missing values in the State ID column were replaced with:

Unknown
This ensures that records with unavailable state information are retained instead of being removed.

Using Find and Replace:

Select the State ID column.
Press Ctrl + H.
In Find what, enter:
~?
In Replace with, enter:
Unknown
Click Replace All.
In Excel, ? is treated as a wildcard character. Therefore, ~? is used when searching for the actual question-mark character.

Tools & Techniques Used
Microsoft Excel
COUNTIF
AVERAGE
ROUND
Pivot Tables
Find and Replace
Missing-value imputation
Mode/Frequency analysis
Data Cleaning Summary
Column / Task	Missing Value Strategy
Missing-value identification	COUNTIF(range,"?")
Month	Replaced with Sep
Year	Replaced with rounded average year
Smoker	Replaced with most frequent value using Pivot Table
Hospital Tier	Replaced with most frequent tier
City Tier	Replaced with most frequent tier
State ID	Replaced with Unknown
Outcome
After completing the cleaning process:

Missing values were identified across both datasets.
Missing month values were replaced with Sep.
Missing year values were replaced with the rounded average year.
Missing categorical values were replaced using their most frequently occurring values.
Missing State ID values were categorized as Unknown.
The resulting dataset is cleaner and more suitable for further analysis and visualization.
Dataset Tables
The analysis primarily uses:

Medical Examinations
Hospitalization Details
These cleaned tables can now be used for additional exploratory data analysis, statistical analysis, dashboard creation, and visualization.

Data Transformation
Overview
After completing the data-cleaning stage, the next step of the project focused on transforming the cleaned data into a more structured and analysis-ready format.

The transformation process was performed using Microsoft Excel and Power Query.

The main objectives were to:

Split customer names into meaningful components
Convert surgery information into numerical data
Standardize inconsistent categorical values
Categorize customers based on BMI and HbA1C
Create Date of Birth and Age fields
Format healthcare charges as currency
Data Transformation Tasks
1. Split Customer Names
The names column in the Customer Names table was split into three meaningful columns:

Title
First Name
Last Name
For example:

Hawks, Ms. Kelly
was transformed into:

Title	First Name	Last Name
Ms.	Kelly	Hawks
Splitting the customer names into separate fields improved data organization and made it easier to filter, sort, and analyze customer information.

2. Convert NumberOfMajorSurgeries to Numerical Data
The NumberOfMajorSurgeries column in the Medical Examinations table contained values that were not consistently stored as numeric data.

Non-numeric characters and textual representations were reviewed and converted into meaningful numerical values.

The objective was to ensure that the column contained only numeric values so that mathematical and analytical operations such as the following could be performed correctly:

Sum
Average
Comparison
PivotTable analysis
After transformation, the NumberOfMajorSurgeries field was converted to a numerical data type.

3. Correct Inconsistencies in Heart Issues and smoker
The Heart Issues and smoker columns were checked for inconsistent categorical values.

Examples of inconsistencies included:

yes
Yes
YES
and:

no
No
NO
Although these values have the same meaning, inconsistent capitalization can cause them to appear as separate categories during analysis.

To resolve this issue, the values were standardized to:

Yes
No
This ensured consistent grouping, filtering, and reporting in PivotTables and dashboards.

4. Create Weight Status
A new column named Weight Status was created based on each customer's BMI value.

The following classification was used:

BMI	Weight Status
Below 18.5	Underweight
18.5 – 24.9	Normal Weight
25.0 – 29.9	Overweight
30.0 and above	Obesity
The following Power Query conditional logic was used:

if [BMI] >= 30 then "Obesity"
else if [BMI] >= 25 then "Overweight"
else if [BMI] >= 18.5 then "Normal Weight"
else "Underweight"
This transformation converted raw BMI values into meaningful health categories that could be used more easily in analysis and visualization.

5. Create Diabetes Status
A new column named Diabetes Status was created using the customer's HBA1C value.

The following classification was used:

HbA1C	Diabetes Status
Below 5.7	Normal
5.7 – 6.4	Prediabetes
6.5 and above	Diabetes
The following Power Query logic was used:

if [HBA1C] >= 6.5 then "Diabetes"
else if [HBA1C] >= 5.7 then "Prediabetes"
else "Normal"
This created a categorical field that made it easier to analyze customers according to their diabetes status.

6. Create Date of Birth
The year, month, and date columns in the Hospitalization Details table were combined into a single column named:

Date of Birth
The resulting date was formatted using the custom format:

DD-MMM-YYYY
Example:

09-Jul-1992
Combining the three separate date components into one field improved readability and allowed the date to be used directly for further calculations, including customer age.

7. Calculate Customer Age
A new column named Age was created using each customer's Date of Birth.

The age was calculated relative to the dataset collection date:

08-Jun-2023
The calculation also considers whether the customer's birthday had already occurred during 2023.

The following Power Query calculation was used:

let
    DOB = Date.FromText([Date of Birth]),
    CollectionDate = #date(2023, 6, 8)
in
    Date.Year(CollectionDate) - Date.Year(DOB) -
    (
        if Date.Month(CollectionDate) < Date.Month(DOB)
        or
        (
            Date.Month(CollectionDate) = Date.Month(DOB)
            and Date.Day(CollectionDate) < Date.Day(DOB)
        )
        then 1
        else 0
    )
This produces each customer's completed age as of 8 June 2023.

Using a fixed collection date instead of the current date ensures that the calculated ages remain consistent with the original dataset.

8. Format Healthcare Charges
The charges column in the Hospitalization Details table was formatted as currency using the US Dollar ($) symbol.

Example:

$13,529.92
This improved readability and ensured that healthcare costs were displayed consistently throughout the analysis and dashboard.

Tools and Techniques Used
The following Excel and Power Query features were used during the transformation stage:

Microsoft Excel
Power Query
Split Column
Conditional Columns
Data Type Conversion
Text Standardization
Date Transformation
Custom Date Formatting
Age Calculation
Currency Formatting
Conditional Logic
Transformation Summary
Transformation	Result
Customer Names	Split into Title, First Name, and Last Name
NumberOfMajorSurgeries	Converted to numerical values
Heart Issues	Standardized to consistent Yes/No values
smoker	Standardized to consistent Yes/No values
BMI	Converted into Weight Status categories
HbA1C	Converted into Diabetes Status categories
year + month + date	Combined into Date of Birth
Date of Birth	Used to calculate Age as of 08-Jun-2023
charges	Formatted as US Dollar currency
New Fields Created
The transformation stage produced several new fields that were later used in the analysis:

Title
First Name
Last Name
Weight Status
Diabetes Status
Date of Birth
Age
These transformed variables were important for creating the final consolidated healthcare dataset and supporting further analysis and dashboard development.

Outcome
After completing the data-transformation process:

Customer names were separated into structured fields.
Major surgery information was converted into numerical format.
Categorical inconsistencies were standardized.
BMI values were categorized into meaningful weight-status groups.
HbA1C values were categorized into diabetes-status groups.
Separate date components were combined into a valid Date of Birth.
Customer ages were calculated as of 8 June 2023.
Healthcare charges were formatted as currency.
The transformed datasets were then ready for the next stage of the project: data consolidation, analysis, and dashboard development.

Data Exploration, Analysis & Visualization
Overview
After data cleaning and transformation, the three source tables were combined into a single Healthcare dataset using Customer ID as the common key.

Microsoft Excel was used to perform:

Data consolidation with VLOOKUP
PivotTable analysis
Pie, Combo, Column, and Scatter charts
Trendline and R² analysis
Interactive dashboard creation with slicers
1. Create the Healthcare Dataset
A new worksheet named Healthcare was created by combining:

Customer Names
Medical Examinations
Hospitalization Details
Customer ID was used as the lookup key.

Example:

=VLOOKUP(A2,'Hospitalisation solutions (2)'!A:B,2,FALSE)
The final Healthcare dataset included fields such as:

Customer ID
First Name
BMI
HBA1C
Heart Issues
Any Transplants
Cancer History
NumberOfMajorSurgeries
smoker
Weight Status
Diabetes Status
Date of Birth
charges
Hospital tier
City tier
State ID
Age
This consolidated table was used for all further analysis.

2. Cancer History Among Smokers and Non-Smokers
A PivotTable was created to compare cancer history between smokers and non-smokers.

Pivot Area	Field
Rows	smoker
Columns	Cancer history
Values	Count of Customer ID
Two Pie Charts were created:

Cancer History Among Smokers
Cancer History Among Non-Smokers
This helped compare the proportion of customers with and without a cancer history across smoking groups.

3. Major Surgeries and HbA1C by Transplant History
A PivotTable was used to compare transplant history with:

Total Number of Major Surgeries
Average HbA1C
Pivot Area	Field
Rows	Any Transplants
Values	Sum of NumberOfMajorSurgeries
Values	Average of HBA1C
A Combo Chart was used:

Major Surgeries → Column
Average HbA1C → Line
HbA1C → Secondary Axis
This made it easier to compare two measures with different numerical scales.

4. Healthcare Charges by Weight and Diabetes Status
Average healthcare charges were compared across:

Weight Status
Diabetes Status
Pivot Area	Field
Rows	Weight Status
Columns	Diabetes Status
Values	Average of charges
A Clustered Column Chart was used to identify which combinations were associated with higher average healthcare costs.

5. Average Charges by Hospital Tier and State
A PivotTable was created to compare average healthcare charges across different states and hospital tiers.

Pivot Area	Field
Rows	State ID
Columns	Hospital tier
Values	Average of charges
A Clustered Column Chart was used for side-by-side comparison.

6. Scatter Plot Analysis
Scatter plots were created to study relationships between numerical variables.

Age vs BMI
X-axis = Age
Y-axis = BMI
Age vs HBA1C
X-axis = Age
Y-axis = HBA1C
Age vs Healthcare Charges
X-axis = Age
Y-axis = charges
A linear trendline and R² value were added to each chart to evaluate the strength and direction of the relationship.

7. Healthcare Analytics Dashboard
All visualizations were combined into an interactive dashboard titled:

Healthcare Analytics Dashboard
The dashboard included:

Cancer History by Smoking Status
Major Surgeries and HbA1C by Transplant History
Charges by Weight Status and Diabetes Status
Charges by Hospital Tier and State
Age vs BMI
Age vs HBA1C
Age vs Healthcare Charges
Two slicers were added:

Weight Status
Diabetes Status
The slicers were connected to supported PivotTables and PivotCharts using Report Connections, allowing users to interactively filter the analysis.

Analysis Summary
Analysis	Visualization
Cancer history by smoking status	Pie Charts
Major surgeries and HbA1C by transplant history	Combo Chart
Charges by Weight and Diabetes Status	Clustered Column
Charges by Hospital Tier and State	Clustered Column
Age vs BMI	Scatter Plot
Age vs HBA1C	Scatter Plot
Age vs Charges	Scatter Plot
Tools & Techniques Used
Microsoft Excel
VLOOKUP
PivotTables
PivotCharts
Pie Charts
Combo Charts
Clustered Column Charts
Scatter Plots
Trendlines
R² Values
Secondary Axis
Currency Formatting
Slicers
Report Connections
Dashboard Design
Final Outcome
The analysis stage produced:

A consolidated Healthcare dataset
Multiple comparative healthcare analyses
Correlation analysis using scatter plots
Healthcare cost comparisons
An interactive Excel dashboard
Dynamic filtering using Weight Status and Diabetes Status slicers
