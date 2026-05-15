# Student-Result-Management-System
Interactive Excel dashboard for student performance analysis. This project helps manage student academic records, calculate results automatically, analyze performance, and visualize insights using dynamic dashboards.

#=> Project Overview

This project is designed to manage and analyze student academic performance in an easy and efficient way.  
This project helps calculate student marks, percentage, grades, and performance automatically using Excel formulas, Pivot Tables, charts, 
and slicers to create an automated result system and interactive dashboard.

#=> Features

-  Automatic Result Calculation
-  Percentage & Grade Generation
-  Pass / Fail Status
-  Student Ranking System
-  Interactive Dashboard
-  Charts & Visual Reports
-  Subject-wise Average Analysis
-  Top 5 Students Section
-  Class & Section Filtering using Slicers

#=> Tools Used

- Microsoft Excel
- Pivot Tables
- Pivot Charts
- Slicers
- VLOOKUP
- IF Formula
- COUNTIF
- Conditional Formatting

#=> Formulas Used

🔹 Total Marks [=SUM(C3:G3)]
🔹 Percentage [ =AVERAGE(C3:G3) ]
🔹 Grade Formula [ =IF(I3>=90,"AA",IF(I3>=80,"A+",IF(I3>=70,"A",IF(I3>=60,"B+",IF(I3>=50,"B",IF(I3>=40,"C+",IF(I3>=30,"C",IF(I3<=30,"Fail")))))))) ]
🔹 Pass / Fail Formula [ =IF(I3>50,"PASS","FAIL")  ]
🔹 Rank Formula [ =RANK(I3,$I$3:$I$32) ]

#=> Dashboard Includes

- KPI Cards
- Total Students
- Average Percentage
- Pass Students
- Highest Percentage
- Student Performance Chart
- Pass vs Fail Chart
- Subject-wise Average Table
- Top 5 Students Table
- Interactive Slicers

#=> Screenshots
1. Dashboard Preview
![image alt](https://github.com/rabinmahato6/Student-Result-Management-System/blob/ec988a5c684a91f35dadb8b6e34d73340978bea7/Dashboard%20Preview.png)

2. Result Sheet
![image alt](https://github.com/rabinmahato6/Student-Result-Management-System/blob/ba036e2063ce73b36f69e5d397080c80dfec9b7f/Result%20Sheet.png)

3. top5 students result
![image alt](

#=> Project Structure
1st create a folder
#Student-Result-Management-System/

──> Student_Result_System.xlsx
──> README.md
──> screenshots/
    ├── dashboard.png
    ├── result_sheet.png
    |── top5 students result.png


#=> Skills Demonstrated

- Data Analysis
- Excel Dashboard Design
- Data Visualization
- Report Generation
- Logical Formula Building

#=> Learning Outcome

Through this project, I improved my knowledge of:
- Advanced Excel formulas
- Dashboard creation
- Data analysis techniques
- Interactive reporting using slicers and charts

 #=> Author

RABINDRA NATH MAHATO
