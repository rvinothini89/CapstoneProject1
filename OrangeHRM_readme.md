# Project Overview

This project is a Selenium-Python-based test automation framework developed to execute a variety of test scenarios on a web application. The framework is structured to manage data, configurations, and driver setups efficiently, ensuring organized and streamlined test execution.

I have designed this framework as a hybrid model, combining the strengths of both data-driven and keyword-driven testing approaches. The data-driven aspect is implemented through an Excel file, which is used for reading login details and updating test outcomes within the same file. The keyword-driven aspect is realized by storing all necessary attribute values in a config.yaml file. This hybrid approach leverages the benefits of both frameworks, providing a flexible and powerful testing solution. 

## Dependencies

To run the tests, the following Python packages are required:

Selenium: For browser automation.

webdriver-manager: To manage browser drivers automatically.

pytest: For running test cases.

pytest-html: For generating HTML reports from test executions.

openpyxl: For handling Excel files.

PyYAML: For reading and writing YAML configuration files.

Install the dependencies using pip:

```bash
pip install selenium webdriver-manager pytest pytest-html openpyxl pyyaml
```

## Folder Structure

TestData Folder

1. Excel File: This file act as test case container which contains login information for all the tests and test case status has been updated after execution.

2. config.yaml File: Stores test inputs required for all test scenarios and test data methods. This includes URLs, expected values for assertions, and other configuration data.

3. Data Class file - This is the file which will read data from excel sheet for login credentials, attribute values from config.yaml and perform web page operations. Mainly web page interactions methods are present here. 

TestLocators Folder 

This folder has a file which contains all locator details for elements on the web page. These locators are referenced and used in data class file to interact with web elements.

TestCodes Folder

1. conftest.py: - This file is responsible for initializing the WebDriver and preventing redundant driver instantiation across test cases.

2. Test Files: This is a python file which uses pytest framework. Define individual test functions required for the tests. Test functions access test data folder for few input parameters from config.yaml, attributes and methods  from data class file. This is the file responsible for updating excel sheet with test case status and execution time scenarios.Each test function is written to validate specific functionalities of the web application.

Utilities Folder

This folder has excel functions and yaml functions class files which are responsible for reading and writing from respective file. 

## Running Tests

1. Ensure all dependencies are installed (Python, Selenium, Pytest, etc.).
2. Place your login information and other test data in the Excel file under the TestData folder.
3. Update the config.yaml file with the necessary inputs for your test scenarios.
4. Run the tests using Pytest using the below command, and check the results and status updates in the Excel file.

```bash
pytest -v -s --capture=sys --html=Reports\CapstoneProject1_Test_report.html .\TestCodes\test_OrangeHRM.py
```

## How it works

Data Input:

Login credentials and other test inputs are read from the Excel file and config.yaml file.
These inputs are used throughout the test execution to perform actions like logging in, adding employees, updating details, etc.

Driver Initialization:

The conftest.py file initializes the WebDriver, ensuring it’s done only once per test session to avoid redundancy.

Test Execution:

Each test case is written as a function in the test files.
These functions use the data and locators to interact with the OrangeHRM application and validate the expected results.
Test statuses, such as "Passed" or "Failed," are written back to the Excel file after execution.

Assertions:

The test functions include assertions to compare actual results with expected outcomes.
These assertions ensure that the application behaves as expected under various scenarios.

## Conclusion
This project structure allows for a clean, organized, and efficient way to manage and execute automated tests. By separating data, locators, and test logic, the project is easy to maintain and extend as the application evolves.