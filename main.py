import os
import re
import datetime
import pandas as pd
from playwright.sync_api import sync_playwright, Page, expect
import logging
import multiprocessing
import dotenv

dotenv.load_dotenv()


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
    'INPUT_ERROR': 'E003',
    'CREATE_ERROR' : 'E004',
    'UNKNOWN_ERROR': 'E999',
}

# --- XPaths ---
LOGIN_ID_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[3]/div[1]/input"
LOGIN_PASSWORD_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[3]/div[3]/input"
LOGIN_BUTTON_XPATH = "/html/body/div[1]/div/div/div[1]/form/div/div/div[4]/div[1]/button"
GET_BY_LABEL_XPATH = '//label[contains(text(), "{label}")]/following-sibling::input'

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

# --- Texts ---
ENROLLMENT_DROPDOWN_TEXT = "Enrollment"

# --- Selectors ---
STUDENT_NAME_SELECTOR = "input#sName"
FATHER_NAME_SELECTOR = "input#fatherName"
SURNAME_SELECTOR = "input#surname"
DOB_SELECTOR = "input#dob"
GENDER_SELECTOR = "input[name='rdoGender'][value='{value}']"
BFORM_SELECTOR = "input#Cnic"
GR_SELECTOR = "input#grNumber"
CELL_SELECTOR = "input#cell"
RELIGION_SELECTOR = "select#religionID"
ADDRESS_SELECTOR = "input#address"
IMG_SELECTOR = "input#imgfile"
CREATE_SELECTOR = "input#btn-create"

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

def select_dropdown(page, element, value, error_code, field_name, gr_no: int, Type: str = "text"):
    """Select an option from a dropdown by visible text."""
    try:
        if Type.lower() == "xpath":
            page.click(f"xpath={element}")
            page.click(f"xpath={value}")
        elif Type.lower() == "text":
            page.click(f"text={element}")
            page.click(f"text={value}")
        elif Type.lower() == "selector":
            page.select_option(element, label=value)
    except Exception as e:
        log_error(
            logger, error_code, f'Error in {field_name}: {e} by using Type {Type}', gr_no
        )
        raise

def fill_input(page, element, value, error_code, field_name, gr_no, is_int=False, Type: str = "selector", typing: bool = False):
    """Fill an input field if value is not NaN."""
    try:
        if pd.notna(value):
            fill_val = str(int(value)) if is_int else str(value).capitalize()
            if Type.lower() == "xpath":
                if typing:
                    page.locator(f"xpath={element}").type(fill_val)
                else:
                    page.locator(f"xpath={element}").fill(fill_val)
            elif Type.lower() == "selector":
                if typing:
                    page.wait_for_selector(element)
                    page.type(element, fill_val)
                else:
                    page.wait_for_selector(element)
                    page.fill(element, fill_val)
                    
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)
        raise

def format_date(date) -> dict["day": int, "month": int, "year": int]:
    """Format date to MM/DD/YYYY."""
    date = pd.Timestamp(str(date)).strftime('%d/%m/%Y')
    return {"day": int(date.split("/")[0]), "month": int(date.split("/")[1]), "year": int(date.split("/")[2])}

def format_religion(data) -> str:
    if data == "Islam":
        return "Muslim"
    else:
        return "NonMuslim"

def fill_date(page, element, value, error_code, field_name, gr_no):
    """Fill a date input field."""
    value = format_date(value)
    day = value["day"]
    month = value["month"]
    year = value["year"]
    try:
        page.click(element)

        # 2. Open year/decade selection
        page.click(".datepicker-days .datepicker-switch")   # Switch to months
        page.click(".datepicker-months .datepicker-switch") # Switch to years
        page.click(".datepicker-years .datepicker-switch")  # Switch to decades

        while True:
            decade_text = page.inner_text(".datepicker-decades .datepicker-switch")
            start, end = [int(x) for x in decade_text.split("-")]
            if start <= year <= end:
                break
            if year < start:
                page.click(".datepicker-decades .prev")
            else:
                page.click(".datepicker-decades .next")

        # 4. Click the decade span to open years
        page.click(f".datepicker-decades .decade:text('{year // 10 * 10}')")

        # 5. Click the year
        page.click(f".datepicker-years .year:text('{year}')")

        # 6. Pick the month
        month_names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"]
        page.click(f".datepicker-months .month:text('{month_names[month-1]}')")

        # 7. Pick the day
        page.click(f".datepicker-days td.day:text('{day}')")
    except Exception as e:
        log_error(logger, error_code, f'Error in {field_name}: {e}', gr_no)
        raise

def fill_radio(page, element, value, error_code, field_name, gr_no):
    """Fill a radio input field."""
    try:
        if pd.notna(value):
            page.check(element.format(value=value))
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
                IMG_SELECTOR,
                image_path,
            )
        else:
            log_error(logger, error_code, f"Image not found for GR {gr_no}", gr_no)
            raise FileNotFoundError(f"Image not found for GR {gr_no}")
    except Exception as e:
        log_error(logger, error_code, f'Error in Image: {e}', gr_no)
        raise FileNotFoundError(f"Image not found for GR {gr_no}")

def _fill_form_sync(data, Username: str, Password: str):
    gr = None
    data.columns = data.columns.str.strip()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()
        page.goto("https://online.bisempk.edu.pk/")

        # --- Login ---
        fill_input(page, LOGIN_ID_XPATH, Username, ERROR_CODES['LOGIN_FAILED'], 'Username', None, Type="xpath")
        fill_input(page, LOGIN_PASSWORD_XPATH, Password, ERROR_CODES['LOGIN_FAILED'], 'Password', None, Type="xpath")
        page.click(f"xpath={LOGIN_BUTTON_XPATH}")
        page.wait_for_timeout(1000)

        
        # --- Iterate Through Excel Rows ---
        len1 = len(data)
        for index, row in data.iterrows():
            try:
                gr = row["GR"]
                # --- Enrollment ---
                select_dropdown(page, ENROLLMENT_DROPDOWN_TEXT, f"SSC-I {row['Group']}", ERROR_CODES['NAVIGATION_FAILED'], 'Enrollment Dropdown', gr)
                page.wait_for_timeout(1000)

                # --- Student Name ---
                fill_input(page, STUDENT_NAME_SELECTOR, row['Student Name'], ERROR_CODES['INPUT_ERROR'], 'Student Name', gr)

                # --- Father Name ---
                fill_input(page, FATHER_NAME_SELECTOR, row['Father Name'], ERROR_CODES["INPUT_ERROR"], 'Father Name', gr)

                # --- Surname ---
                fill_input(page, SURNAME_SELECTOR, row['Surname'], ERROR_CODES["INPUT_ERROR"], 'Surname', gr)

                # --- DOB ---
                fill_date(page, DOB_SELECTOR, row['D.O.B'], ERROR_CODES["INPUT_ERROR"], 'D.O.B', gr)

                # --- Gender ---
                fill_radio(page, GENDER_SELECTOR, row['Gender'], ERROR_CODES["INPUT_ERROR"], 'Gender', gr)
                
                # --- B Form ---
                fill_input(page, BFORM_SELECTOR, row['B.Form No'], ERROR_CODES["INPUT_ERROR"], 'B-Form', gr, is_int=True, typing=True) 

                # --- GR Number ---
                fill_input(page, GR_SELECTOR, row['GR'], ERROR_CODES["INPUT_ERROR"], 'GR', gr, is_int=True)

                # --- Cell ---
                fill_input(page, CELL_SELECTOR, f"0{row['Mobile No.']}", ERROR_CODES["INPUT_ERROR"], 'Cell', gr)

                # --- Religion ---
                select_dropdown(page, RELIGION_SELECTOR, format_religion(row['Religion']), ERROR_CODES["INPUT_ERROR"], 'Religion', gr, Type="Selector")

                # --- Address ---
                fill_input(page, ADDRESS_SELECTOR, row['Address'], ERROR_CODES["INPUT_ERROR"], "Address", gr)

                # --- Image ---
                upload_image(page, gr, ERROR_CODES["INPUT_ERROR"])

                # --- Create ---
                # try:
                #     page.click(CREATE_SELECTOR)
                # except Exception as e:
                #     log_error(logger, ERROR_CODES["CREATE_ERROR"], f'Error in Create: {e}', gr)
                #     raise

                # ---
                page.wait_for_timeout(20000)
                page.goto("https://online.bisempk.edu.pk/")
                page.wait_for_timeout(2000)
            except Exception as e:
                log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e}", gr)
                continue  # Proceed to next row

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
        data = pd.read_excel("data.xlsx")
        fill_form_from_excel(data, dotenv.getenv("Username"),  dotenv.getenv("Password"))
    except Exception as e:
        log_error(logger, ERROR_CODES['UNKNOWN_ERROR'], f"Error: {e}", None)