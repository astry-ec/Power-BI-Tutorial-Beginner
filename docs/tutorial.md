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
d. You will notice the data is loaded in the Data panel, in the right side. We call the loaded data as Recruitment_Data table.
<p>&nbsp;</p>
<img width="249" height="523" alt="image" src="https://github.com/user-attachments/assets/150af89d-3109-42e4-989d-cbee3e273a8c" />
<p>&nbsp;</p>

### 3. Create your first DAX measures

Power BI uses DAX (Data Analysis Expressions), a formula language designed for data analysis. Some values required for our analysis are not directly available in the raw data. Therefore, we need to perform calculations to get the required values.

In this tutorial, we will learn basic DAX calculations that we need to build our dashboard.

Go to **Modeling** -> **New Measure**. We will create 3 measures for this dashboard. You can directly copy the formula below into "measure" pane

#### a. Total Applicants
```text
Total Applicants = DISTINCTCOUNT(Recruitment_Data[ApplicationID])
```
<p>&nbsp;</p>
<img width="900" height="183" alt="image" src="https://github.com/user-attachments/assets/88b7bd1d-2eb2-464e-a555-de23b2b0bca0" />
<p>&nbsp;</p>

#### b. Total Hired

```text
Total Hired = 
CALCULATE (
    DISTINCTCOUNT ( 'Recruitment_Data'[ApplicationID] ),
    'Recruitment_Data'[Status] = "Hired"
)
```
<p>&nbsp;</p>
<img width="896" height="251" alt="image" src="https://github.com/user-attachments/assets/fc9a1ab5-7f7f-4b0a-9e2a-d7f3ff5f3437" />
<p>&nbsp;</p>

#### c. Hire Rate

```text
Hire Rate = DIVIDE([Total Hired], [Total Applicants])
```
<p>&nbsp;</p>
<img width="876" height="187" alt="image" src="https://github.com/user-attachments/assets/08246b1f-36d6-4ef7-a811-779e774e94d6" />
<p>&nbsp;</p>
After creating all measures data, we can find the measures in Recruitment_Data Table
<p>&nbsp;</p>
<img width="341" height="464" alt="image" src="https://github.com/user-attachments/assets/9767ddbe-34d6-4529-a3aa-aac1413546c8" />
<p>&nbsp;</p>

### 4. Create Visualizations
There are a lot of chart types provided by Power BI. In this tutorial, We will create 5 visualizations. 
<p>&nbsp;</p>
#### 1. Card Visualizations
A card displays only a number, which is a best practice for showing a single key metric. Each card displays a pre-defined measure. For this example, we will use measure : Total Applicants, Total Hired and Hire Rate.
<p>&nbsp;</p>
full documentation of Card Visualizations : https://learn.microsoft.com/en-us/power-bi/visuals/power-bi-visualization-card
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="454" height="649" alt="image" src="https://github.com/user-attachments/assets/650910a2-43ed-4e53-8b23-078f73bf9b70" />
<p>&nbsp;</p>
- Drag Total Applicants, Total Hired, and Hire Rate to Values field.
<p>&nbsp;</p>
<img width="1060" height="619" alt="image" src="https://github.com/user-attachments/assets/d29d0584-44a9-4bde-bfc6-638cb356d550" />
<p>&nbsp;</p>
- Card visual appears with the 3 measures added.
<p>&nbsp;</p>
#### 2. Column Chart Visualizations
Column charts display data as vertical bars that can be grouped by category. Power BI offers several types of column charts. In this tutorial, we will use a *Clustered column chart* to display the Total Applicants and Total Hired Employees by Job Category.
Full documentation of Column Chart : https://learn.microsoft.com/en-us/power-bi/paginated-reports/report-design/visualizations/column-charts-report-builder
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="360" height="507" alt="image" src="https://github.com/user-attachments/assets/17932787-746d-4236-ba68-0e04731f82ca" />
<p>&nbsp;</p>
- Drag Job_Title to X-Axis. Then, Total Applicants and Total Hire to Y-Axis.
<p>&nbsp;</p>
<img width="1184" height="490" alt="image" src="https://github.com/user-attachments/assets/4a6b9240-5051-4062-91cd-c61bd8e5db99" />
<p>&nbsp;</p>
#### 3. Pie Chart Visualizations
<p>&nbsp;</p>
A pie chart displays data as a proportion of a metric. In this tutorial, we will use the proportion of total applicants to job vacancy sources as the metric.
<p>&nbsp;</p>
Full documentation of Pie Chart : https://learn.microsoft.com/en-us/power-bi/paginated-reports/report-design/visualizations/pie-charts-report-builder
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="362" height="524" alt="image" src="https://github.com/user-attachments/assets/5fab2fb8-a5ed-4da5-82de-0cf5c5fb7e90" />
<p>&nbsp;</p>
- Drag Source_Name to Legend field. Then, Total Applicants to Values field.
<p>&nbsp;</p>
<img width="842" height="493" alt="image" src="https://github.com/user-attachments/assets/be169867-8329-439e-b583-58e782842743" />
<p>&nbsp;</p>
#### 4. Line Chart Visualizations
A line chart displays data changes over time. In this tutorial, we will use data on the total number of applicants of each month.
<p>&nbsp;</p>
Full documentation of Line Chart : https://learn.microsoft.com/en-us/power-bi/paginated-reports/report-design/visualizations/line-charts-report-builder.
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="371" height="527" alt="image" src="https://github.com/user-attachments/assets/1b02069d-aa25-40a6-bc07-766850952a5c" />
<p>&nbsp;</p>
- Drag StatusDate to X-axis and Total Applicants to Y-axis.
<p>&nbsp;</p>
<img width="988" height="555" alt="image" src="https://github.com/user-attachments/assets/a3999ef9-a02e-4257-91d5-8cd78b5b820f" />
<p>&nbsp;</p>
- Power BI will automatically break down StatusDate into Year, Quarter, Month, and Day. The line chart will display the highest level of the StatusDate hierarchy: Year. Since our data contains only the year 2026, the line chart shows just a single point. To change the X-axis to months, we need to remove Year, Quarter, and Day from the X-axis by clicking the "x" icon. The chart will change the display as below
<p>&nbsp;</p>
<img width="991" height="495" alt="image" src="https://github.com/user-attachments/assets/ff14fec9-744f-4734-aef6-d4e3b9569ae4" />
<p>&nbsp;</p>
#### 5. Table Visualizations
A table visualization is the most straightforward and flexible visual. It looks exactly like a standard table consisting of columns and rows, and, any necessary columns can be added. 
Full table visualizations: https://learn.microsoft.com/en-us/power-bi/visuals/power-bi-visualization-tables?tabs=powerbi-desktop
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="371" height="513" alt="image" src="https://github.com/user-attachments/assets/77a34fa4-1924-4c6c-9dab-bd7a66de8886" />
<p>&nbsp;</p>
- Drag the data field as shown in picture below
<p>&nbsp;</p>
<img width="976" height="511" alt="image" src="https://github.com/user-attachments/assets/223d6cb4-bbba-41dd-b189-6a30413dd629" />
<p>&nbsp;</p>
### 5. Format Page and Visualizations
We have created visualizations, and to make the dashboard more engaging and dynamic, we will explore the formatting features in Power BI.
In this part, we will learn how to format the dashboard page and its visualizations. In addition, we will also learn to add slicers to make dashboard filtering easier.
<p>&nbsp;</p>
#### 1. Format Page
<p>&nbsp;</p>
The dashboard page has default aspect ration of 16:9 (1920 x 1080). However, if our visualizations require more space for our visualizations, we customize the page dimensions. 

Default size
<img width="401" height="640" alt="image" src="https://github.com/user-attachments/assets/c28ef192-1953-4b06-b2ab-413d5ef82265" />
<p>&nbsp;</p>
Go to **Visualizations -> Format Page -> Canvas Settings -> Type**
<img width="375" height="485" alt="image" src="https://github.com/user-attachments/assets/7c36625d-cafe-4416-8804-4624aedf5975" /> <img width="366" height="492" alt="image" src="https://github.com/user-attachments/assets/ff67dac2-7643-40bd-8450-4488d33ec809" />
<p>&nbsp;</p>
#### 2. Format Visual
Power BI provides various options for formatting visualizations. In this tutorial, we will focus on formatting the legend, title and X-axis and Y-axis labels for Clustered Column Chart.

- Legend
- <p>&nbsp;</p>
There are several options for positioning the legend in a visualizations. In this chart, we select Top right. We can also format the font and color, but for now we will leave them as they are
<img width="1005" height="607" alt="image" src="https://github.com/user-attachments/assets/7095a4a5-d11f-4287-b966-339970d4fdd5" />
<p>&nbsp;</p>
- Title
<p>&nbsp;</p>
To make the chart easier to understand, we need to use clear and readable title. In this chart, we simply remove the underscore (_) from Job_Title in the title and subtitle.
<img width="1008" height="834" alt="image" src="https://github.com/user-attachments/assets/0010f0a5-290d-40ae-9bb2-103160b8f8c6" />
<p>&nbsp;</p>
- X-axis and Y-axis Labels
<p>&nbsp;</p>
In this chart, we will add title for X-axis and change maximum number of Y-axis
<img width="1008" height="657" alt="image" src="https://github.com/user-attachments/assets/bbdd1247-596a-45f8-834d-e605b21d7609" />
<img width="1008" height="557" alt="image" src="https://github.com/user-attachments/assets/4ce7f49c-d9af-4c9c-8f40-6e11776d6ac0" />
<p>&nbsp;</p>

#### 3. Create Slicer
Slicer is one of visualizations which can be functioned as a filter of the page. 
<p>&nbsp;</p>
- Hover your cursor over the visualization icon as shown below, then click it.
<p>&nbsp;</p>
<img width="365" height="516" alt="image" src="https://github.com/user-attachments/assets/d40a3c28-f6a2-4db6-ae93-c247e73cd416" />
<p>&nbsp;</p>
- Drag Source_name to Values field as below
<p>&nbsp;</p>
<img width="715" height="526" alt="image" src="https://github.com/user-attachments/assets/4bf1c081-324b-4841-820e-fa7cfe638cc9" />
<p>&nbsp;</p>
At this point, our dashboard should look like the image below. When no value is selected in the Source Name slicer in the upper-right corner, the dashboard displays data from all sources (Image 1). After we select "JobStreet", the dashboard updates to display only data from the "JobStreet" source (Image 2).
<p>&nbsp;</p>
Image 1
<img width="1318" height="810" alt="image" src="https://github.com/user-attachments/assets/af75f190-35c8-421a-ac22-f9d368a804ea" />
<p>&nbsp;</p>
Image 2
<img width="1323" height="809" alt="image" src="https://github.com/user-attachments/assets/960a9bc6-f57c-43b0-bee7-ed2ef5ba7b24" />
<p>&nbsp;</p>
To clear the filter and display all data again, click "JobStreet" again.

#### 4. Dashboard Page Finishing
After preparing the data, creating the visualizations, and formatting them, we come to the final step: improving the appearance of our dashboard. Here are some best practices we can follow.
<p>&nbsp;</p>
- Add a dashboard title using a text box.
<p>&nbsp;</p>
<img width="1450" height="935" alt="image" src="https://github.com/user-attachments/assets/6b25d686-d325-4831-9df5-f6f8f2b84eed" />
<p>&nbsp;</p>
Select all the text, then choose the desired font size.
<p>&nbsp;</p>
<img width="844" height="231" alt="image" src="https://github.com/user-attachments/assets/3b9aa7b8-149c-4d70-9a31-aebb198e6b93" />
<p>&nbsp;</p>
- Add rectangle shapes to organize and separate the visualizations.
<p>&nbsp;</p>
This step is optional. You can add a background shape to give your dashboard a webpage-like appearance and choose any color you prefer.
<img width="1502" height="932" alt="image" src="https://github.com/user-attachments/assets/dcf76e81-0934-412a-8ab2-5054a148d47e" />
<p>&nbsp;</p>
- Select a theme from the View pane. You can choose any theme you prefer.
<p>&nbsp;</p>
<img width="1530" height="947" alt="image" src="https://github.com/user-attachments/assets/4c528bb1-2027-481d-889b-bb9a14af415a" />
<p>&nbsp;</p>















