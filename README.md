# Paramount Hospital: Surgical Outcomes and Patient Recovery Data Analysis (Interactive Dashboard creation using Microsoft Excel)

![surgery-1807541_1280](https://github.com/user-attachments/assets/b17b7166-1026-4b02-bd8e-33e775f96299)

### Project Overview
Paramount Hospital wants to create an annual report for 2018 - 2020, to analyze the relationships between key patient demographics (Age, BMI) and incontinence symptoms (ICIQ scores), in order to have a clear understanding of the factors influencing surgical outcomes and patient recovery in their medical operations and services.


### Data Sources
The primary dataset used for this analysis is the Combined data_No_pt_info.xlsx file, with specific instructions to work only with the BULKAMID Sheet in the Excel file. This sheet contains detailed information about each aspects of the surgical operations both pre and post. 

- <a href="https://github.com/ezekielokpara/data-analysis-with-excel/addmilesales.xlsx">Download Data Set Here"</a>

^ copy the url of your file from the github project space, after uploading it to github, and paste it as seen above.


### Exploratory Data Analysis (Questions)

EDA involved explorying and analyzing BulkAmid data to answer key questions, such as:

- What is the Total Number of Suggeries performed in a given year/year range and quarter?
- What is the Total Number of Suggeries performed in a given month/month range?
- What is the Number of Successful and Unsuccessful Suggeries performed in a given year/year range?
- How many patients who come in for surgery have Prolapse, and how many don't, in a given year/year range?
- Calculate and determine the Average and Standard Deviation of Patients Age, BMI, ICIQ Pre, ICIQ Post and Post OP FU (Monthly)
- What is the Average BMI by Age Group?
- Determine the effect of suggery on patients. Did more patients improved or not?
- What is the overall BMI analysis for patients?
- What is the effect of patients BMI on ICIQ Pre?
- What is the effect of patients BMI on ICIQ Post?
- How does a patients Age affect his/her BMI?
- What is effect of Patients Age on Surgical Outcomes - (How does a patients Age affect ICIQ Post)?
- What is the Percentage of Complications at FU?
- What is the percentage of None-Complications at FU 
- Represent Parity Values and groups with a chart.
- Dashboard interaction [View Dashboard](#)


### Tools
- Excel 
-- Power Query
-- Pivot Table Analyze
-- Excel Data Analysis Functions 
-- Excel Formulars
-- Charts
-- Dashboard Creation Utilities for creating Reports & Visualization
-- Data Cleaning [Download Here](https://microsoft.com)

### Data Cleaning and Standardization

In the initial data preparation phase, I performed the following tasks:

1. Data loadaing and inspection
2. Verification of data for any missing values and anomalies.
3. Loaded data into Excel Power Query
4. In Excel Power Query, I performed the following tasks:
- Handled missing values and anomalies.
- Data transformation
- Standardized Data Types for various columns fields
- Carried out Data Cleaning and formating. This is to make sure data is consistent, and clean with respect to the data type, data format and values used.
- Loaded transformed data into a New Excel Worksheet for analysis


### Preparing Data for Analysis

1. Created Pivot tables according the data analysis questions asked.

2. Merged all Pivot tables into dashboards and applied Slicers to make data filtering dynamic and interactive.


### Data Analysis

#### Descriptive Statistics

- I calculated basic Descriptive Statistics (Average (mean) and Standard Deviation) for key variables like age, BMI, and ICIQ scores to understand the general characteristics of the dataset.

- Age Distribution: Included a section summarizing the age range (30-35, 35-40, 40-45, 45-50, 50-55, 55-60, 60-65, 65-70, 70-75, 80-85) of patients and the Average BMI of each Age Group.

#### Correlation Analysis

- Using the Correllation (CORREL) Function in Excel, I explored and analyzed the relationships between different variables, such as BMI, Age, and incontinence severity.

-- To acheive this, I ran a Correlation Analysis for:
--- **BMI and ICIQ pre-op**: This will tell if higher BMI is associated with worse incontinence symptoms before surgery.
--- **BMI and ICIQ post-op**: This will tell if BMI is associated with post-op incontinence outcomes.
--- **Age and ICIQ pre-op**: This will tell if older age is correlated with more severe pre-operative incontinence symptoms.
--- **Age and BMI**: This will help determine if older patients tend to have higher BMI values.

- Here are the standard **BMI (Body Mass Index)** categories and their corresponding ranges, as defined by the **World Health Organization (WHO)** used in this analysis:

Table goes here.
| Correllation Value | Relationship Status |
|--------------------|---------------------|
|
|

These BMI categories help determine whether a person has a healthy weight in relation to their height, and they are often used in healthcare to assess risks related to weight, such as cardiovascular diseases, diabetes, and more.

- Using the Correllation (CORREL) Function in Excel, I also explored and analyzed whether age impacts surgery success rates by comparing age groups with post-operative outcomes such as symptom reduction.

- I used the following table values to determine the Correllation or relationship between these factors:

| Correllation Value | Relationship Status |
|--------------------|---------------------|
|
|

#### Comparative Analysis

- By Comparative Analysis, I evaluate the effectiveness of surgeries by looking at the difference between Average ICIQ pre-op vs. Average ICIQ post-op data. I checked if the Average ICIQ pre-op was lower than the Average ICIQ post-op data to determine if more MORE Patients improved after suggery or not.


### Findings and Results

1. More Patients improved after Suggery
2. With an average BMI value of 25, More Patients who come for suggery are Overweight
3. The Correlation or relationship between BMI and ICIQ Pre is Weak Negative
4. There is no Correlation between BMI and Age (BMI is not affected by Age)
5. There is no Correlation between Age and ICIQ Post (Age does not impact surgical outcomes)
6. There are more number (93%) of Complications at FU and about 7% without Complications at FU.
7. There are more suggeries the 3rd Quarter (July, August, September) of each year.
8. The lowest number of suggeries fall within the 1st & 2nd Quarter of the Year (January - June), with the 1st Quarter recording the lowest.

Add DASHBORD IMAGE HERE


### Recommendations

- **Improvement Areas**: Based on the analysis, I suggest that more focus should be deployed to provide intensive follow-up care after operations. This is due to the fact that from the analysis, we have a high percentage of complications arising after suggery. 

-- Though, based on analysed data which shows that more patients improve after surgery, for there to be complications arising after surgery, it means there are issues with the follow-up protocol in place after surgery. 

-- Potential areas for improving patient care, should be in providing better treatment and follow-up protocols and future patient care strategies. This should be done by targeting patients with specific characteristics for certain types of surgery, or adjusting treatment plans based on pre-operative conditions.

- **Treatment Efficacy**: Generally, even with incidents of complications arising after surgery, it is noted from the analysis that patients do improve after suggery. It is therefore, recommended that the underlying cause of the complications be investigated and dealt with to move the process of suggery to near perfect. Better and improved surgcial methods can be reserched and adopted as well.

- Overall, the efficacy of the treatments and surgery offered, based on the data analysed, seems particular effective. 
