import os
import re
import datetime
import pandas as pd
from playwright.sync_api import sync_playwright, Page, expect
import logging
import multiprocessing

date = datetime.datetime.now().strftime("%d-%m-%Y")
# --- Logging Setup ---
logging.basicConfig(
    filename=f'logs/{date}.txt',
    level=logging.ERROR,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

ERROR_CODES = {
    'LOGIN_FAILED': 'E001',
    'NAVIGATION_FAILED': 'E002',
    'UNKNOWN_ERROR': 'E999',
}

# --- XPaths ---
LOGIN_ID_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[3]/div[1]/input"
LOGIN_PASSWORD_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[3]/div[3]/input"
LOGIN_BUTTON_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[4]/div[1]/button"
ENROLLMENT_DROPDOWN_XPATH = "/html/body/div[1]/div/div[1]/div[2]/ul/li[5]/a"
ENROLLMENT_OPTION_XPATH = "/html/body/div[1]/div/div[1]/div[2]/ul/li[5]/ul/li[1]/a"
MAT_OPTION_SPAN_XPATH = "//mat-option/span[normalize-space(text())='{value}']"
MAT_LABEL_XPATH = "//mat-label[contains(., '{field_name}')]"
MAT_SELECT_XPATH = "ancestor::mat-form-field//mat-select"
IMAGE_UPLOAD_INPUT_XPATH = "/html/body/app-root/app-main-layout/div/app-add-student/section/div/div[2]/div/div/div[2]/form/div[2]/div/div/div[2]/div[2]/div[1]/app-file-upload/div/input"
STUDENT_ADMISSION_BUTTON_XPATH = "/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/a"
ADD_STUDENT_BUTTON_XPATH = "/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[2]/a"
STUDENT_LIST_BUTTON_XPATH = "/html/body/app-root/app-main-layout/app-sidebar/div/aside/div/ul/li[5]/ul/li[1]/a"
SEARCH_INPUT_XPATH = "/html/body/app-root/app-main-layout/div/app-all-students/section/div/div[2]/div/div/div/div/div/div/div[1]/div/div[1]/ul/li[2]/input"
SEARCH_BUTTON_XPATH = "/html/body/app-root/app-main-layout/div/app-all-students/section/div/div[2]/div/div/div/div/div/div/div[1]/div/div[1]/ul/li[3]/div/button"
MAT_ROW_XPATH = "ancestor::mat-row"
STATUS_TAB_XPATH = "/html/body/app-root/app-main-layout/div/app-about-student/section/div/div[2]/div[3]/div/mat-tab-group/mat-tab-header/div/div/div/div[3]/div"
CHANGE_STATUS_BUTTON_XPATH = "/html/body/app-root/app-main-layout/div/app-about-student/section/div/div[2]/div[3]/div/mat-tab-group/div/mat-tab-body[3]/div/div/div/student-academic-year-status/div/div[1]/div[2]/div/button"


def log_error(logger, error_code, message, gr_no=None):
    """Log errors with consistent format and speak them."""
    error_msg = f"[{error_code}] {message}"
    if gr_no:
        error_msg += f" (GR NO: {gr_no})"
    logger.error(error_msg)
    return error_msg


# --- Helper Functions ---
def select_mat_option_by_label(page: Page, label: str, value: str):
    # Find the mat-select combobox by its accessible label (e.g. "Items per page:")
    combobox = page.get_by_role("combobox", name=label)

    # Open the dropdown
    combobox.click()

    # Angular Material renders the options inside a listbox overlay
    listbox = page.get_by_role("listbox").first
    expect(listbox).to_be_visible()

    # Locate the option by its exact visible text
    option = listbox.get_by_role("option", name=value)

    # Scroll in case it's off-screen and click
    option.scroll_into_view_if_needed()
    option.click()

    # Verify the combobox now shows the chosen value
    expect(combobox).to_have_text(value)


def select_dropdown(page, element, value, error_code, field_name, gr_no, Type="ID"):
    """Select an option from a dropdown by visible text."""
    try:
        if Type == "ID":
            page.click(element)
            page.click(value)
        # elif Type == "TEXT":
        #     status_label = page.locator(
        #         f"xpath={MAT_LABEL_XPATH.format(field_name=field_name)}"
        #     )
        #     status_select = status_label.locator(f"xpath={MAT_SELECT_XPATH}")
        #     status_select.click()
        #     page.click(f"xpath={MAT_OPTION_SPAN_XPATH.format(value=value)}")
    except Exception as e:
        log_error(
            logger, error_code, f'Error in {field_name}: {e} by using Type {Type}', gr_no
        )
        raise


def fill_input(page, xpath, value, error_code, field_name, gr_no, is_int=False):
    """Fill an input field if value is not NaN."""
    try:
        if pd.notna(value):
            fill_val = str(int(value)) if is_int else str(value)
            page.locator(f"xpath={xpath}").fill(fill_val)
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)
        raise


def fill_date(page, xpath, value, error_code, field_name, gr_no):
    """Fill a date input field with formatted date."""
    try:
        if pd.notna(value):
            formatted_date = pd.Timestamp(str(value)).strftime('%m/%d/%Y')
            page.locator(f"xpath={xpath}").fill(formatted_date)
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)
        raise


def upload_image(page, gr_no, error_code):
    """Upload image if exists for the given GR NO."""
    try:
        image_dir = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "Photos"
        )
        image_path = os.path.join(image_dir, f"{gr_no}.jpg")
        if os.path.exists(image_path):
            page.set_input_files(
                f"xpath={IMAGE_UPLOAD_INPUT_XPATH}",
                image_path,
            )
        else:
            log_error(logger, error_code, f"Image not found for GR {gr_no}", gr_no)
            raise FileNotFoundError(f"Image not found for GR {gr_no}")
    except Exception as e:
        log_error(logger, error_code, f'Error in Image: {e}', gr_no)
        raise


def navigate_to(page, option: str, error_code, gr_no):
    """Navigate to a specific section in the EMIS portal."""
    try:
        page.click(f"xpath={STUDENT_ADMISSION_BUTTON_XPATH}")
        if option == "Add Student":
            page.click(f"xpath={ADD_STUDENT_BUTTON_XPATH}")
            page.wait_for_timeout(2000)
        elif option == "Student List":
            page.click(f"xpath={STUDENT_LIST_BUTTON_XPATH}")
            page.wait_for_timeout(1000)
    except Exception as e:
        log_error(logger, error_code, f"Error in navigation: {e}", gr_no)
        raise


def select_student_by_gr(page, gr_no, error_code):
    """Select a student from the list by GR NO."""
    try:
        # navigate to student list
        navigate_to(page, "Student List", ERROR_CODES['NAVIGATION_FAILED'], gr_no)
        select_mat_option_by_label(page, "Items per page:", "100")  # Show 100 entries
        fill_input(
            page,
            SEARCH_INPUT_XPATH,
            gr_no,
            ERROR_CODES['INPUT_ERROR'],
            "Search",
            gr_no,
        )
        page.wait_for_timeout(100)
        page.click(f"xpath={SEARCH_BUTTON_XPATH}")
        page.wait_for_timeout(1000)
        # Build exact regex pattern
        pattern = re.compile(rf"^\s*{str(gr_no)}\s*$")

        # Find the GR cell with exact match
        gr_cell = page.locator("mat-cell.cdk-column-grNo").filter(has_text=pattern)

        student = gr_cell.locator(f"xpath={MAT_ROW_XPATH}")
        student.locator("button[mat-icon-button]").nth(0).click()
    except Exception as e:
        log_error(
            logger, error_code, f"Error selecting student with GR NO {gr_no}: {e}", gr_no
        )
        raise


def Go_to_edit_Status(page, error_code, gr_no):
    """Navigate to the Status tab of the selected student."""
    try:
        page.click(f"xpath={STATUS_TAB_XPATH}")
        page.wait_for_timeout(1000)

        # Click 'Change Status' button
        page.click(f"xpath={CHANGE_STATUS_BUTTON_XPATH}")
        page.wait_for_timeout(1000)
    except Exception as e:
        log_error(logger, error_code, f"Error navigating to Status tab: {e}", gr_no)
        raise


def _fill_form_sync(data, Username: str, Password: str):
    ver = None
    data.columns = data.columns.str.strip()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://online.bisempk.edu.pk/")

        # --- Login ---
        fill_input(page, LOGIN_ID_XPATH, Username, ERROR_CODES['LOGIN_FAILED'], 'Username', None)
        fill_input(page, LOGIN_PASSWORD_XPATH, Password, ERROR_CODES['LOGIN_FAILED'], 'Password', None)
        page.click(f"xpath={LOGIN_BUTTON_XPATH}")
        page.wait_for_timeout(1000)

        # --- Enrollment ---
        select_dropdown(page, f"xpath={ENROLLMENT_DROPDOWN_XPATH}", f"xpath={ENROLLMENT_OPTION_XPATH}", ERROR_CODES['NAVIGATION_FAILED'], 'Enrollment', None)
        page.wait_for_timeout(2000)
        # --- Iterate Through Excel Rows ---
        # len1 = len(data)
        # for index, row in data.iterrows():
        #     ver = row["GR NO"]

        browser.close()


# --- Main Form Filling Logic ---
def fill_form_from_excel(data, Username, Password):
    process = multiprocessing.Process(
        target=_fill_form_sync, args=(data, Username, Password)
    )
    process.start()
    process.join()


# --- Entry Point ---
if __name__ == '__main__':
    try:
        multiprocessing.freeze_support()
        fill_form_from_excel(pd.read_excel("data.xlsx"), "340@bisempk.edu.pk", "Sailboat20667=")
    except Exception as e:
        log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e}", None)