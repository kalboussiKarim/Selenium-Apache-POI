# Selenium-Apache-POI

## Overview
This repository provides a comprehensive guide and 
practical examples for interacting with Excel documents 
using Apache POI in Java. The project focuses on 
enabling seamless automation workflows where data-driven 
testing plays a crucial role. It includes methods to 
**read from**, **write to**, and **update Excel files**, 
along with a real-world scenario demonstrating the 
integration of Apache POI with Selenium for effective 
test result management. Additionally, it includes 
utilities to interact with properties files for better 
configuration management.

## Features
### 1. **Working with Excel Files**
- **Writing Data**: Includes methods to write data to specific rows and cells in an Excel file.
- **Reading Data**: Implements utilities to read data from Excel sheets, including handling edge cases like empty rows or cells.
- **Updating Excel Files**: Demonstrates how to update specific cells with dynamic data, such as test results (`Pass`/`Fail`).

### 2. **Real-World Testing Scenario**
- A complete example is provided where:
    1. Test data is fetched from an Excel sheet.
    2. The expected result is compared with the actual result.
    3. The Excel sheet is updated with the test status (`Pass`/`Fail`).
- This showcases how Apache POI integrates seamlessly w
- ith Selenium to enable data-driven testing.

## Technologies Used
- ![Java](https://img.shields.io/badge/Java-%23ED8B00.svg?style=for-the-badge&logo=java&logoColor=white) **Java**: Core language for implementing the utilities and examples.
- ![Apache POI](https://img.shields.io/badge/Apache%20POI-%23A81C7D.svg?style=for-the-badge&logo=apache&logoColor=white) **Apache POI**: For Excel file interactions (reading, writing, updating).
- ![Selenium](https://img.shields.io/badge/Selenium-%2343B02A.svg?style=for-the-badge&logo=selenium&logoColor=white) **Selenium**: For web automation testing.

## How to Use
1. Clone the repository:
   ```bash
   git clone https://github.com/kalboussiKarim/Selenium-Apache-POI.git

2. Import the project into your favorite IDE (e.g., IntelliJ IDEA, Eclipse).
3. Update the test_data folder with your Excel files and properties files.
4. Run the test examples provided in the src/test/java folder.

## Examples

### Example: Writing Data to an Excel File
```java
ExcelUtilities.setCellData("TestData.xlsx", "Sheet1", 1, 1, "Test Passed");
```

### Example: Reading Data from an Excel File
```java
String data = ExcelUtilities.getCellData("TestData.xlsx", "Sheet1", 1, 1);
System.out.println("Data from Excel: " + data);
```