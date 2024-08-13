from selenium.common import TimeoutException, NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains, Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from TestLocators.OrangeHrm_Locators import locators
from Utilities.excel_functions import excelFunction
from Utilities.yaml_functions import YAMLReader


class OrangeHRMData:
    # Class to handle OrangeHRM operations with test data and interactions

    # config file which are needed for reading input and writing output
    excel_file = "D:\\VinoLEarning\\Capstone_Vino\\TestData\\OrangeHRM_TestData.xlsx"
    yaml_file = "D:\\VinoLEarning\\Capstone_Vino\\TestData\\config.yaml"
    sheet_number = "Sheet1"

    # Initialize data readers for Excel and YAML
    excel_obj = excelFunction(excel_file, sheet_number)
    yaml_obj = YAMLReader(yaml_file)

    # Read configuration and test data from YAML
    url = yaml_obj.reader()['url']
    dashboard_url = yaml_obj.reader()['dashboard_url']
    title = yaml_obj.reader()['title']
    searchstring = yaml_obj.reader()['employee_searchstring']
    error_text = yaml_obj.reader()['error_text']
    pass_data = yaml_obj.reader()['pass_data']
    fail_data = yaml_obj.reader()['fail_data']

    def __init__(self, url, driver):
        """
        Initialize with URL and WebDriver instance.
        """
        self.driver = driver
        self.url = url
        self.wait = WebDriverWait(driver, 30)
        self.excel_obj = excelFunction(self.excel_file, self.sheet_number)

    @classmethod
    def read_login_data(cls):
        """
        Read login data for testing from the Excel file.
        Returns a list of tuples containing username, password, and row number.
        """
        data = []
        # range 2,4 is used as login scenarios are present in this range in excel sheet
        for row in range(2, 4):
            # column number 6 and 7 has username and password values in excel sheet respectively
            username = cls.excel_obj.read_data(row, 6)
            password = cls.excel_obj.read_data(row, 7)
            data.append((username, password, row))
        return data

    # created another method to read data for employee operations like add, update and delete.
    @classmethod
    def read_login_data_empOperations(cls, start_row, end_row=None):
        """
        Read login data for employee operations from the Excel file.
        If end_row is not provided, use start_row as the only row.
        """
        data = []
        end_row = end_row or start_row
        for row in range(start_row, end_row + 1):
            username = cls.excel_obj.read_data(row, 6)
            password = cls.excel_obj.read_data(row, 7)
            data.append((username, password, row))
        return data

    def WebPageAccess(self):
        """
        Access the specified URL and maximize the browser window.
        Returns the title of the current page.
        """
        try:
            self.driver.maximize_window()
            self.driver.get(self.url)
            print(self.driver.title)
            return self.driver.title
        except TimeoutException as e:
            print(e)

    def login(self, username, password):
        """
        Perform login with provided username and password.
        Handles both successful and failed login attempts.
        """
        try:
            username_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_username)))
            username_element.send_keys(username)
            print(username)
            pswd_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_password)))
            pswd_element.send_keys(password)
            print(password)
            login_button = self.wait.until(EC.element_to_be_clickable((By.XPATH, locators.loc_login)))
            login_button.click()
            # Check for error message indicating a failed login
            try:
                error_message_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_error)))
                print(error_message_element.text)
                error_text = error_message_element.text
                return {"status": "failure", "error_message": error_text}
            except TimeoutException:
                current_url = self.driver.current_url
                return {"status": "success", "url": current_url}
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def PIMAccess(self):
        """
        Access the PIM section of the application.
        """
        try:
            PIM_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_PIM)))
            PIM_element.click()
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def ClickAdd(self):
        """
        Click the 'Add' button to initiate employee addition.
        """
        try:
            Add_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_AddButton)))
            Add_element.click()
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def AddEmployeeDetails(self):
        """
        Add employee details including first name, last name, and employee ID.
        """
        try:
            firstname_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_firstname)))
            firstname_value = self.yaml_obj.reader()['first_name']
            firstname_element.send_keys(firstname_value)
            lastname_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_lastname)))
            lastname_value = self.yaml_obj.reader()['last_name']
            lastname_element.send_keys(lastname_value)
            employeeId_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_employeeid)))
            self.driver.execute_script("arguments[0].value = '';", employeeId_element)
            employeeID_value = self.yaml_obj.reader()['employeeID']
            employeeId_element.send_keys(employeeID_value)
            savebutton_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_savebutton)))
            savebutton_element.click()
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def AddPersonalDetailsPart1(self):
        """
        Add personal details including 'Other ID'.
        Waits for success notification to disappear before proceeding.
        """
        try:
            self.wait.until(EC.url_contains("/viewPersonalDetails"))
            print(self.driver.current_url)
            # Wait until the success notification disappears
            self.wait.until(EC.invisibility_of_element((By.XPATH, locators.loc_success)))
            # Locate and interact with the 'Other ID' field
            otherID_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, locators.loc_otherID)))
            otherID_value = self.yaml_obj.reader()['otherID']
            self.driver.execute_script("arguments[0].value = arguments[1];", otherID_element, otherID_value)
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def AddPersonalDetailsPart2(self):
        """
        Add additional personal details including driving license, license expiry date, nation, marital status, date of birth, and gender.
        """
        try:
            # Locate and interact with the driving license field
            license_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_driving_license)))
            licensenum_value = self.yaml_obj.reader()['license_number']
            self.driver.execute_script("arguments[0].value = arguments[1];", license_element, licensenum_value)

            # Locate and interact with the license expiry date field
            licenseExpiry_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_licenseexpirydate)))
            licenseExpiry_value = self.yaml_obj.reader()['license_expiry']
            licenseExpiry_element.send_keys(licenseExpiry_value)

            # Locate and interact with the nation dropdown
            nation_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_nationSelect))) and self.wait.until(EC.visibility_of_element_located((By.XPATH, locators.loc_nationSelect)))
            nation_element.click()
            # this is the control with dropdown and having dynamic values, so used keys down to select input
            for _ in range(10):  # Adjust the number of keys as needed
                nation_element.send_keys(Keys.DOWN)
            nation_element.click()

            # Locate and interact with the marital status dropdown
            marital_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_maritalStatus)))
            marital_element.click()
            marital_element.send_keys(Keys.DOWN)
            marital_element.click()

            # Locate and interact with the date of birth field
            dob_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_dob)))
            dob_value = self.yaml_obj.reader()['dob']
            dob_element.send_keys(dob_value)

            # Locate and interact with the gender radio button
            gender_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_femaleGender)))
            gender_element.click()

            # Click the save button and wait for the success message
            savebutton_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_savebutton)))
            savebutton_element.click()
            # Waiting for success message appear and disappear
            self.wait.until(EC.visibility_of_element_located((By.XPATH, locators.loc_success)))
            self.wait.until(EC.invisibility_of_element((By.XPATH, locators.loc_success)))
            return True

        except TimeoutException as e:
            print(e)
            return False

        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def CheckCreatedUser(self):
        try:
            employeelist_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_employeeList)))
            employeelist_element.click()
            self.empSearch(self.searchstring)
            editbutton_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_editButton)))
            editbutton_element.click()
            self.wait.until(EC.url_contains("/viewPersonalDetails"))
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def empSearch(self, searchstring):
        """
        Search for an employee by name.
        """
        try:
            self.wait.until(EC.url_contains('/viewEmployeeList'))
            empsearch_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, locators.loc_empsearch)))
            empsearch_element.click()
            actions = ActionChains(self.driver)
            searchstring = self.yaml_obj.reader()['employee_searchstring']
            actions.move_to_element(empsearch_element).click().send_keys(searchstring).perform()
            searchbutton_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_searchButton)))
            searchbutton_element.click()
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def modifyEmployeeDetails(self):
        """
        Modify employee details including 'Other ID' and license expiry date.
        """
        try:
            editbutton_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_editButton)))
            editbutton_element.click()
            self.wait.until(EC.url_contains("/viewPersonalDetails"))
            otherID_element = self.wait.until(EC.element_to_be_clickable((By.XPATH, locators.loc_otherID)))
            otherID_value = self.yaml_obj.reader()['otherID']
            self.driver.execute_script("arguments[0].value = 'otherID_value';", otherID_element)
            licenseExpiry_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_licenseexpirydate)))
            licenseExpiry_value = self.yaml_obj.reader()['license_expiry_update']
            licenseExpiry_element.send_keys(licenseExpiry_value)
            savebutton_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_savebutton)))
            savebutton_element.click()
            self.wait.until(EC.visibility_of_element_located((By.XPATH, locators.loc_success)))
            self.wait.until(EC.invisibility_of_element((By.XPATH, locators.loc_success)))
            return True
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False

    def deleteEmployeeDetails(self):
        """
        Delete employee details and confirm deletion.
        """
        try:
            deletebutton_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_deleteButton)))
            deletebutton_element.click()
            confirm_element = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_confirm)))
            confirm_element.click()
            # Wait for success message to disappear
            self.wait.until(EC.invisibility_of_element((By.XPATH, locators.loc_success)))
            searchbutton_element = self.wait.until(
                EC.presence_of_element_located((By.XPATH, locators.loc_searchButton)))
            searchbutton_element.click()
            search_result = self.wait.until(EC.presence_of_element_located((By.XPATH, locators.loc_search_result)))
            search_result_text = search_result.text
            return search_result.text
        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False
        except NoSuchElementException as e:
            print(f"NoSuchElementException occurred: {e}")
            return False
