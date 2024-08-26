---
layout: default
---

## 🛠️ Tools and Technologies Used

![Python](https://img.shields.io/badge/Python-%2314354C.svg?style=for-the-badge&logo=python&logoColor=white)
![Power BI](https://img.shields.io/badge/Power_BI-%23F2C811.svg?style=for-the-badge&logo=power-bi&logoColor=black)
![EasyOCR](https://img.shields.io/badge/OCR-%23000000.svg?style=for-the-badge&logo=OCR&logoColor=white)
![DAX](https://img.shields.io/badge/DAX-%23007ACC.svg?style=for-the-badge&logo=powerbi&logoColor=white)
![Pandas](https://img.shields.io/badge/Pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)
![Matplotlib](https://img.shields.io/badge/Matplotlib-%23ffffff.svg?style=for-the-badge&logo=Matplotlib&logoColor=black)
![Seaborn](https://img.shields.io/badge/Seaborn-%23001a72.svg?style=for-the-badge&logo=seaborn&logoColor=white)
![Regular Expressions](https://img.shields.io/badge/Regex-%23e34f26.svg?style=for-the-badge&logo=Regex&logoColor=white)
![DAX](https://img.shields.io/badge/DAX-%23007ACC.svg?style=for-the-badge&logo=powerbi&logoColor=white)




## 📊 Exploratory Data Analysis (EDA)

This project involves extracting, processing, and analyzing data from images of receipts using Optical Character Recognition (OCR) and Python. The key steps are:

- **Text Extraction**: Using EasyOCR to read text from images.
- **Data Cleaning**: Using Regular Expressions to identify and clean relevant data such as dates and numerical values.
- **Data Transformation**: Formatting extracted data into structured tables with Pandas.
- **Data Visualization**: Using Matplotlib and Seaborn for plotting data distributions and trends.
- **Excel Integration**: Automating the process of adding structured data to Excel files using Openpyxl.

## 🧠 Key Analysis Techniques

| **Technique**           | **Description**                                                                                  |
|-------------------------|--------------------------------------------------------------------------------------------------|
| **OCR Text Extraction** | Leveraging EasyOCR to extract text from image files, converting it into readable formats.        |
| **Dax Data Parsing**  | Applying Regular Expressions to identify patterns, extract relevant data, and handle text cleanup.|
| **Excel Automation**    | Using Openpyxl to automate the process of updating Excel sheets with new data entries.           |

![Flowchart diagram](./assets/flowchart_diagram.png)

## 📈 Key Findings

1. **Efficient Text Extraction**: EasyOCR was effective in reading and extracting relevant text from receipts, demonstrating its capability in processing various fonts and layouts.
2. **Data Structuring**: The use of Pandas allowed for efficient structuring and manipulation of data into a tabular format suitable for analysis and storage.
3. **Automation**: Integrating Openpyxl for Excel automation streamlined the process of updating records, making the workflow more efficient and less prone to errors.

## 📊 Power BI Report

To enhance the analysis and visualization capabilities, Power BI was employed to create detailed and interactive reports. The key features of the Power BI reports include:

- **Data Integration**: Connecting processed data from Python scripts directly to Power BI for real-time updates.
- **Advanced Visualizations**: Utilizing Power BI’s rich set of visualizations to provide insights into spending patterns.
- **DAX Functions**: Implementing DAX to create complex measures and calculated columns for more in-depth analysis.

![Weekly Expense Report](./assets/report.png)
