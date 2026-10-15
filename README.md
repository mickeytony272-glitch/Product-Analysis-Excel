#Excel Data Exploration and Analysis

Project Overview

This project focuses on performing basic data exploration and analysis using Microsoft Excel. The objective is to analyze a product dataset using Excel formulas and functions to summarize data, apply conditional logic, calculate category-based metrics, and extract information from text fields.

This assignment demonstrates foundational Excel skills that are essential for data cleaning, data exploration, and preparation for further analysis.

Dataset Description

The dataset contains product-level information with the following attributes:

Column

Description

Product ID

Unique identifier for each product

Product Name

Name of the product

Brand Name

Brand associated with the product

Price ($)

Price of the product

Quantity

Quantity of the product

Category

Product category

Price Range

Classification based on product price

Day

First two characters of Product ID

Country Code

Last two characters of Product ID

Month

Characters extracted from the Product ID

Dataset: Product Dataset
Tool Used: Microsoft Excel
Dataset Size: 34 product records

Objectives

Explore and summarize product data.

Calculate total, average, minimum, and maximum prices.

Categorize products using logical functions.

Perform conditional calculations based on product categories and prices.

Extract information from Product IDs using text functions.

Tasks Performed

1. Basic Data Exploration

Used Excel functions to calculate key summary statistics.

Analysis

Excel Function

Formula

Total Product Price

SUM

=SUM(D2:D35)

Number of Products

COUNTA

=COUNTA(B2:B35)

Average Price

AVERAGE

=AVERAGE(D2:D35)

Minimum Price

MIN

=MIN(D2:D35)

Maximum Price

MAX

=MAX(D2:D35)

2. Price Categorization Using IF

Created a new column named Price Range to classify products based on their price.

Business Rules:

Price greater than or equal to $500: High Price

Price less than $500: Standard Price

Formula:

=IF(D2>=500,"High Price","Standard Price")

The formula can be filled down to categorize every product in the dataset.

3. Conditional Analysis Using SUMIF and COUNTIF

Used conditional functions to answer specific business questions.

Total price of Electronics products

=SUMIF(F2:F35,"Electronics",D2:D35)

This calculates the sum of prices for products belonging to the Electronics category.

Number of products priced below $100

=COUNTIF(D2:D35,"<100")

This counts the products whose prices are less than $100.

4. Text Extraction Using LEFT, RIGHT, and MID

Extracted specific characters from the Product ID column to create three additional columns.

New Column

Excel Function

Formula

Purpose

Day

LEFT

=LEFT(A2,2)

Extracts the first two characters

Country Code

RIGHT

=RIGHT(A2,2)

Extracts the last two characters

Month

MID

=MID(A2,4,3)

Extracts three characters starting from the fourth position

These formulas assume that Product IDs follow a consistent text format.

Key Skills Demonstrated

Microsoft Excel data exploration

Basic statistical calculations

Logical functions and conditional categorization

Conditional aggregation and counting

Text manipulation and extraction

Formula application across multiple records

Data summarization and interpretation

Tools and Technologies

Microsoft Excel

Excel formulas and functions: SUM, COUNTA, AVERAGE, MIN, MAX, IF, SUMIF, COUNTIF, LEFT, RIGHT, and MID

GitHub for project documentation and version control

Key Learnings

Through this project, I practiced using Excel to summarize datasets, classify records based on business rules, calculate category-specific metrics, and extract structured information from text fields.

These skills provide a foundation for more advanced data analysis using Excel, SQL, Python, and business intelligence tools.

Conclusion

This assignment demonstrates how Microsoft Excel can be used to perform fundamental data analysis tasks efficiently. By combining mathematical, logical, conditional, and text functions, the project transforms raw product data into more structured and useful information.

This project is part of my learning journey toward building practical data analytics skills and developing a portfolio of data analysis projects.
