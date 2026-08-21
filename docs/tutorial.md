# Build Your First Recruitment Dashboard in Power BI

## What you will build
In this tutorial, we will learn step by step how to build Power BI Dashboard with Excel file as the data source. Along the way we will also learn how to create DAX measures and calculated columns to get the required KPI data.
In addition, we will learn how to build some charts using several types of Power BI chart and creating slicer to narrow down the data.

## Prerequisites
Before we begin this tutorial, make sure you have:

- Power BI Desktop installed on your computer.
- The Excel file that will be used as the data source throughout this tutorial.
- Basic familiarity with Power BI, such as navigating the Power BI Desktop interface and using its main features.

### 1. Download the sample dataset
   
- Go to the data folder
- Locate the "tutorial recruitment data.xlsx" file
- Download the file

### 2. Import the data into Power BI

a. From "Home", find the "Get Data" and click Excel Workbook
<p>&nbsp;</p>
<img width="445" height="570" alt="image" src="https://github.com/user-attachments/assets/27e187e7-085e-4fd6-8a81-ca896ce71632" />
<p>&nbsp;</p>
b. Find the "tutorial recruitment data.xlsx" in your local folder
<p>&nbsp;</p>
<img width="756" height="480" alt="image" src="https://github.com/user-attachments/assets/644afc18-c1fa-461c-aa44-a4124ea362f9" />
<p>&nbsp;</p>
c. The file will be loaded in Navigator pane as below. We can click the table name and check the data inside the table. Tick table "Recruitment Data".
  We don't need to change any data from the source, we just need to click "Load".
<p>&nbsp;</p>
<img width="899" height="710" alt="image" src="https://github.com/user-attachments/assets/a607da0e-40cd-4b38-9c90-c87e1a2730d8" />
<p>&nbsp;</p>
d. You will notice the data is loaded in the Data panel, in the right side.
<p>&nbsp;</p>
<img width="249" height="523" alt="image" src="https://github.com/user-attachments/assets/150af89d-3109-42e4-989d-cbee3e273a8c" />
<p>&nbsp;</p>

### 3. Create your first DAX measures

Power BI uses DAX (Data Analysis Expressions), a formula language designed for data analysis. Some values required for our analysis are not directly available in the raw data. Therefore, we need to perform calculations to get the required values.

In this tutorial, we will learn basic DAX calculations that we need to build our dashboard.
Go to **Modeling** -> **New Measure**
<p>&nbsp;</p>
a. Total Applicants
```text
Total Applicants = DISTINCTCOUNT(Recruitment_Data[ApplicationID])
```
<p>&nbsp;</p>
<img width="900" height="183" alt="image" src="https://github.com/user-attachments/assets/88b7bd1d-2eb2-464e-a555-de23b2b0bca0" />

