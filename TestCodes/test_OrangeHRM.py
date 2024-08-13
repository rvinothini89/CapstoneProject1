import pytest
from selenium.common import TimeoutException
from TestData.OrangeHrm_Data import OrangeHRMData
from Utilities.excel_functions import excelFunction
from Utilities.yaml_functions import YAMLReader


class Test_OrangeHRM:
    """
        using pytest fixture for initiating driver per class and driver details are provided in conftest.py file
        conftest.py is used as i am initiating chrome driver using it and i am avoiding multiple instantiation of driver
        using this.
    """

    @pytest.fixture(autouse=True)
    # this setup class will be run for every test function as the scope is function
    # using setup class for creating object for OrangeHRMData class to access all its methods
    def setup_class(self, driver):
        self.driver = driver
        """
        seems to be defect with pytest as simple instantiation object thrown error, as per pytest community, initializing
        object with type keyword
        """
        type(self).obj = OrangeHRMData(OrangeHRMData.url, driver)
        self.obj.WebPageAccess()
        type(self).eobj = excelFunction(OrangeHRMData.excel_file, OrangeHRMData.sheet_number)
        type(self).yobj = YAMLReader(OrangeHRMData.yaml_file)

    # defined this method as login is needed for each test and this method is called in all the tests
    def login_helper(self, username, password):
        try:
            result = self.obj.login(username, password)
            if result['status'] == "success":
                # Assert that the URL matches the expected success URL
                assert result['url'] == OrangeHRMData.dashboard_url, f"Expected success URL does not match. Got {result['url']}"
                print(f"Successfully logged in using {username}")
                return True
            else:
                # Handle unexpected result
                print(f"Login failed {result}")
                return False

        except TimeoutException as e:
            print(f"TimeoutException occurred: {e}")
            return False

    # Test responsible for testing both successful and invalid logins
    # parametrizing the test so that tests will be repeated for all possible input
    @pytest.mark.parametrize("username,password,row", OrangeHRMData.read_login_data())
    def test_login(self, username, password, row):
        try:
            # Read expected error text and success URL
            expected_error_text = self.yobj.reader()['error_text']
            expected_success_url = OrangeHRMData.dashboard_url

            # Perform login and capture the result
            result = self.obj.login(username, password)

            # Check the status of the result
            if result['status'] == "success":
                # Assert that the URL matches the dashboard url post login
                assert result[
                           'url'] == expected_success_url, f"Expected success URL does not match. Got {result['url']}"
                print(f"Successfully logged in using {username}")
                self.eobj.write_data(row, 8, OrangeHRMData.pass_data)

            elif result['status'] == "failure":
                # Assert that the error text matches the expected error text
                assert result[
                           'error_message'] == expected_error_text, f"Expected error text does not match. Got {result['error_message']}"
                print(f"Login failed as expected with error: {result['error_message']}")
                self.eobj.write_data(row, 8, OrangeHRMData.fail_data, result['error_message'])

            else:
                # Handle unexpected result
                raise AssertionError(f"Unexpected result: {result}")

        except AssertionError as e:
            # Handle assertion failure (unexpected result)
            print(f"AssertionError occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))

        except TimeoutException as e:
            # Handle timeout exceptions specifically
            print(f"TimeoutException occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))

        except Exception as e:
            # Handle any other unexpected exceptions
            print(f"Unexpected error occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))

    # Test responsible for testing add new employee behavior
    @pytest.mark.parametrize("username,password,row", OrangeHRMData.read_login_data_empOperations(4))
    def test_AddEmployee(self, username, password, row):
        try:
            # Attempt login and check if successful
            login_status = self.login_helper(username, password)
            assert login_status, "Login was not successful"

            # Perform employee addition operations
            self.obj.PIMAccess()
            self.obj.ClickAdd()
            self.obj.AddEmployeeDetails()
            self.obj.AddPersonalDetailsPart1()
            self.obj.AddPersonalDetailsPart2()

            # Check if adding personal details part 2 was successful
            create_status = self.obj.CheckCreatedUser()
            assert create_status, F"Expected True, but got {create_status}"

            # Update status in Excel
            print("Employee got added successfully")
            self.eobj.write_data(row, 8, OrangeHRMData.pass_data)

        except AssertionError as e:
            # Handle assertion errors and update Excel status
            print(f"AssertionError occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed

        except TimeoutException as e:
            # Handle timeout exceptions and update Excel status
            print(f"TimeoutException occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed

    # Test responsible for testing modifying an employee behavior
    @pytest.mark.parametrize("username,password,row", OrangeHRMData.read_login_data_empOperations(5))
    def test_ModifyEmployee(self, username, password, row):
        try:
            # Attempt login and check if successful
            login_status = self.login_helper(username, password)
            assert login_status, "Login was not successful"

            # Perform employee details updation
            self.obj.PIMAccess()
            self.obj.empSearch(OrangeHRMData.searchstring)
            success_update_status = self.obj.modifyEmployeeDetails()
            assert success_update_status, "Failed to modify an existing employee details"
            print("Employee details modified successfully")
            self.eobj.write_data(row, 8, OrangeHRMData.pass_data)

        except AssertionError as e:
            # Handle assertion errors and update Excel status
            print(f"AssertionError occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed

        except TimeoutException as e:
            # Handle timeout exceptions and update Excel status
            print(f"TimeoutException occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed

    # Test responsible for deleting an employee behavior
    @pytest.mark.parametrize("username,password,row", OrangeHRMData.read_login_data_empOperations(6))
    def test_DeleteEmployee(self, username, password, row):
        try:
            # Attempt login and check if successful
            login_status = self.login_helper(username, password)
            assert login_status, "Login was not successful"

            expected_delete_status = self.yobj.reader()['expected_delete_text']

            # Perform employee details deletion
            self.obj.PIMAccess()
            self.obj.empSearch(OrangeHRMData.searchstring)
            actual_delete_status = self.obj.deleteEmployeeDetails()
            assert actual_delete_status == expected_delete_status, "Failed to delete an existing employee details"
            print("Employee details deleted successfully")
            self.eobj.write_data(row, 8, OrangeHRMData.pass_data)

        except AssertionError as e:
            # Handle assertion errors and update Excel status
            print(f"AssertionError occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed

        except TimeoutException as e:
            # Handle timeout exceptions and update Excel status
            print(f"TimeoutException occurred: {e}")
            self.eobj.write_data(row, 8, OrangeHRMData.fail_data, str(e))
            pytest.fail(str(e))  # Mark the test as failed
