class locators:
    # Locators for the login page
    loc_username = "//input[@name='username']"  # Username input field
    loc_password = "//input[@name='password']"  # Password input field
    loc_login = "//button[text()=' Login ']"  # Login button
    loc_error = "//p[@class = 'oxd-text oxd-text--p oxd-alert-content-text']"  # Error message for login failures
    loc_forgotpswd = "//p[@class = 'oxd-text oxd-text--p orangehrm-login-forgot-header' and text() = 'Forgot your password? ']"  # Forgot password link

    # Locators for the password reset screen
    loc_reset_username = "//input[@name='username' or @placeholder='Username']"  # Username input field on reset screen
    loc_reset_button = "//button[@class='oxd-button oxd-button--large oxd-button--secondary orangehrm-forgot-password-button orangehrm-forgot-password-button--reset']"  # Reset button
    loc_reset_msg = "//h6[@class='oxd-text oxd-text--h6 orangehrm-forgot-password-title']"  # Success message after password reset

    # Locators for the main page
    loc_PIM = "//span[text() = 'PIM']"  # PIM (Personnel Information Management) tab

    # Locators for adding a new employee
    loc_AddButton = "//button//i[@class = 'oxd-icon bi-plus oxd-button-icon']"  # Button to add a new employee
    loc_firstname = "//input[@name = 'firstName']"  # First name input field
    loc_lastname = "//input[@name = 'lastName']"  # Last name input field
    loc_employeeid = "//label[text()='Employee Id']//following::input[@class='oxd-input oxd-input--active']"  # Employee ID input field
    loc_savebutton = "//button[@class='oxd-button oxd-button--medium oxd-button--secondary orangehrm-left-space']"  # Save button after entering employee details
    loc_success = "//div[@class='oxd-toast oxd-toast--success oxd-toast-container--toast']"  # Success message for successful operations
    loc_notification_text = "//p[@class = 'oxd-text oxd-text--p oxd-text--toast-message oxd-toast-content-text']"  # General notification text

    # Locators for employee personal details
    loc_otherID = "(//input[@class='oxd-input oxd-input--active'])[3]"  # Other ID input field
    loc_driving_license = "(//label[text()=\"Driver's License Number\"]//following::input[@class='oxd-input oxd-input--active'])[1]"  # Driving license input field
    loc_licenseexpirydate = "(//input[@class = 'oxd-input oxd-input--active'])[5]"  # License expiry date input field
    loc_nationSelect = "(//div[@class = 'oxd-select-text-input'])[1]"  # Nationality dropdown
    loc_maritalStatus = "(//div[@class = 'oxd-select-text-input'])[2]"  # Marital status dropdown
    loc_dropdown = "//div[@class ='oxd-select-text oxd-select-text--focus']"  # Dropdown menu
    loc_dob = "(//label[text()='Date of Birth']//following::input[@class='oxd-input oxd-input--active'])[1]"  # Date of birth input field
    loc_maleGender = "//label[text()='Male']"  # Male gender option
    loc_femaleGender = "//label[text()='Female']"  # Female gender option
    loc_employeeList = "//a[@class = 'oxd-topbar-body-nav-tab-item' and text() = 'Employee List']"

    # Locators for modifying employee details
    loc_tree = "//button[@class='oxd-icon-button oxd-main-menu-button']"  # Tree menu button
    loc_empsearch = "(//div[@class='oxd-autocomplete-text-input oxd-autocomplete-text-input--active'])[1]"  # Employee search field
    loc_searchButton = "//button[@class='oxd-button oxd-button--medium oxd-button--secondary orangehrm-left-space']"  # Search button
    loc_editButton = "//i[@class='oxd-icon bi-pencil-fill']"  # Edit button
    loc_deleteButton = "(//i[@class='oxd-icon bi-trash'])[1]"  # Delete button
    loc_confirm = "//button[@class='oxd-button oxd-button--medium oxd-button--label-danger orangehrm-button-margin']"  # Confirmation button for deletion
    loc_webpage_cont = "//div[@class='orangehrm-container']"  # Container for the main webpage content
    loc_search_result = "//span[@class='oxd-text oxd-text--span' and text()='No Records Found']"  # Message when no records are found

    # Locators for page headers
    loc_usermgmt = "//span[@class='oxd-topbar-body-nav-tab-item' and text()='User Management ']"  # User Management tab
    loc_job = "//span[@class='oxd-topbar-body-nav-tab-item' and text()='Job ']"  # Job tab
    loc_org = "//span[@class='oxd-topbar-body-nav-tab-item' and text()='Organization ']"  # Organization tab
    loc_qual = "//span[@class='oxd-topbar-body-nav-tab-item' and text()='Qualifications ']"  # Qualifications tab
    loc_nation = "//a[@class='oxd-topbar-body-nav-tab-item' and text()='Nationalities']"  # Nationalities tab
    loc_corp_bank = "//a[@class='oxd-topbar-body-nav-tab-item' and text()='Corporate Branding']"  # Corporate Branding tab
    loc_config = "//span[@class='oxd-topbar-body-nav-tab-item' and text()='Configuration ']"  # Configuration tab
    loc_Admin = "//span[text() = 'Admin']"  # Admin tab

    # Locators for tabs in the application
    loc_Admintab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Admin']"  # Admin tab in the sidebar
    loc_Leavetab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Leave']"  # Leave tab in the sidebar
    loc_Timetab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Time']"  # Time tab in the sidebar
    loc_Recruittab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Recruitment']"  # Recruitment tab in the sidebar
    loc_infotab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'My Info']"  # My Info tab in the sidebar
    loc_perftab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Performance']"  # Performance tab in the sidebar
    loc_dashtab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Dashboard']"  # Dashboard tab in the sidebar
    loc_directtab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Directory']"  # Directory tab in the sidebar
    loc_maintab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Maintenance']"  # Maintenance tab in the sidebar
    loc_claimtab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Claim']"  # Claim tab in the sidebar
    loc_buzztab = "//span[@class = 'oxd-text oxd-text--span oxd-main-menu-item--name' and text() = 'Buzz']"  # Buzz tab in the sidebar
