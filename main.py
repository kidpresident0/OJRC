import ctypes
import logging
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager
from selenium import webdriver
from datetime import datetime
import os


# Constants to prevent sleep
#ES_CONTINUOUS = 0x80000000
#ES_SYSTEM_REQUIRED = 0x00000001

# prevent_sleep():
    #ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS | ES_SYSTEM_REQUIRED)

#def allow_sleep():
    #ctypes.windll.kernel32.SetThreadExecutionState(ES_CONTINUOUS)


# Excel Loader
def load_excel(file_path):
    try:
        # Load the Excel file into a DataFrame
        data = pd.read_excel(file_path, engine='openpyxl')
        print("Columns in Excel after loading:", data.columns.tolist())
        return data
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        print(f"Error loading Excel file: {e}")
        return None

# CSV Loader
def load_csv(file_path):
    try:
        data = pd.read_csv(file_path, encoding="utf-8-sig")
        print("Columns in CSV after loading:", data.columns.tolist())
        return data
    except Exception as e:
        logging.error(f"Error loading CSV file: {e}")
        print(f"Error loading CSV file: {e}")
        return None

# File Loader (determines if CSV or Excel)
def load_file(file_path):
    if file_path.lower().endswith('.csv'):
        return load_csv(file_path)
    elif file_path.lower().endswith('.xlsx'):
        return load_excel(file_path)
    else:
        print("Unsupported file format. Please provide a CSV or Excel (.xlsx) file.")
        return None

# Excel Writer
def update_excel(file_path, data, original_columns):
    if not data:
        logging.info("No data to write to Excel.")
        print("No data to write to Excel.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input Excel, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the DataFrame to an Excel file without index
    try:
        df.to_excel(file_path, index=False, engine='xlsxwriter')
        logging.info(f"Data successfully written to Excel: {file_path}")
        print(f"Data successfully written to Excel: {file_path}")
    except Exception as e:
        logging.error(f"Error writing data to Excel: {e}")
        print(f"Error writing data to Excel: {e}")

# CSV Writer
def update_csv(file_path, data, original_columns):
    if not data:
        logging.info("No data to write to CSV.")
        print("No data to write to CSV.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input CSV, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the DataFrame to a CSV file without index
    try:
        df.to_csv(file_path, index=False)
        logging.info(f"Data successfully written to CSV: {file_path}")
        print(f"Data successfully written to CSV: {file_path}")
    except Exception as e:
        logging.error(f"Error writing data to CSV: {e}")
        print(f"Error writing data to CSV: {e}")



def init_webdriver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # This runs Chrome in headless mode
    options.add_experimental_option("detach", True)
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    return driver


def search_gdc(driver, doc_number, first_name, last_name):
    driver.get("https://docpub.state.or.us/OOS/searchCriteria.jsf")

    # Wait for the page to load fully
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "body")))

    # Check for "Agree" button and click it if present
    try:
        agree_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.ID, "disclaimerForm:btnAgree"))
        )
        agree_button.click()
    except TimeoutException:
        pass  # No need to log the absence or presence of the 'Agree' button every time.

    # Remove leading zeros from DOC number
    doc_number = str(int(doc_number))  # Convert to int and back to string to remove leading zeros

    # Proceed with searching by DOC number in the mainBodyForm:SidNumber field
    try:
        # Locate the SID Number input field by its ID
        sid_field = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "mainBodyForm:SidNumber"))
        )
        sid_field.clear()
        sid_field.send_keys(doc_number)
        sid_field.send_keys("\n")  # Press Enter to submit
        time.sleep(2)
    except TimeoutException:
        error_message = f"Error locating the SID Number input field for DOC number {doc_number}, Name: {first_name} {last_name}"
        logging.info(error_message)
        print(error_message)
        return None

    # Click on the DOC number link, which would show the inmate details
    try:
        doc_link = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.LINK_TEXT, doc_number))
        )
        doc_link.click()
    except TimeoutException:
        error_message = f"Error finding AIC with DOC number: {doc_number}, Name: {first_name} {last_name}"
        logging.info(error_message)
        print(error_message)
        return None

    # Extract inmate details
    return extract_inmate_details(driver)

def extract_inmate_details(driver):
    try:
        name = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "offensesForm:name"))
        ).text

        location = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.XPATH,
                                              "//a[@title='The state institution or county where the offender is serving their sentence.']"))
        ).text

        status = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "offensesForm:status"))
        ).text

        release_date = WebDriverWait(driver, 10).until(
            EC.visibility_of_element_located((By.ID, "offensesForm:relDate"))
        ).text

        return {
            "Name": name,
            "Location": location,
            "Status": status,
            "Release Date": release_date
        }

    except Exception as e:
        print(f"Error extracting inmate details: {e}")
        return None


def process_individual(driver, doc_number, first_name, last_name, stop_flag):
    if stop_flag():
        logging.info("Stopping individual processing as requested.")
        return None

    # Call search_gdc with first name and last name
    result = search_gdc(driver, doc_number, first_name, last_name)
    if result:
        # Add DOC number, first name, and last name to the result
        result["DOCNumber"] = doc_number
        result["firstName"] = first_name
        result["lastName"] = last_name

        # Log successful processing
        success_message = f"Successfully processed DOC number: {doc_number}, Name: {first_name} {last_name}"
        logging.info(success_message)
        print(success_message)  # Print to console to show activity

        return result
    else:
        return None




def process_with_retries(index, doc_number, first_name, last_name, stop_flag):
    if stop_flag():
        logging.info("Stopping process with retries as requested.")
        return index, None

    retries = 3
    for attempt in range(retries):
        if stop_flag():
            logging.info("Stopping WebDriver initialization as requested.")
            return index, None

        driver = init_webdriver()
        try:
            result = process_individual(driver, doc_number, first_name, last_name, stop_flag)
            return index, result
        except WebDriverException as e:
            logging.error(f"WebDriver exception on attempt {attempt + 1} for DOC number {doc_number} ({first_name} {last_name}): {e}")
            time.sleep(1)
        finally:
            driver.quit()

    return index, None



def update_csv(file_path, data, original_columns):
    if not data:
        logging.info("No data to write to CSV.")
        print("No data to write to CSV.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input CSV, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the CSV without an index
    try:
        df.to_csv(file_path, index=False)
        logging.info(f"Data successfully written to {file_path} with {len(data)} rows.")
        print(f"Data successfully written to {file_path} with {len(data)} rows.")
    except Exception as e:
        logging.error(f"Error writing data to CSV: {e}")
        print(f"Error writing data to CSV: {e}")



import os
import time
import logging
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

def run_main_process(input_file, output_file, stop_flag):
    start_time = time.time()
    data = load_file(input_file)  # Use load_file to load CSV or Excel
    if data is None or data.empty:
        print("Failed to load data from file.")
        return

    # Clean column names by stripping whitespace and converting to consistent case
    data.columns = data.columns.str.strip().str.lower()

    # Rename columns to match expected names if they are different in input
    column_rename_map = {
        'last name': 'lastname',
        'fist name': 'firstname',
        'doc number': 'docnumber',
        'previous location': 'previouslocation',
        'current': 'current',
        'out?': 'out',
        'street': 'street',
        'city': 'city',
        'state': 'state',
        'zip': 'zip',
        'date of search': 'date of search',
        'initials': 'initials',
        'phone': 'phone',
        'notes': 'notes'
    }
    data.rename(columns=column_rename_map, inplace=True)

    if 'docnumber' not in data.columns:
        print("Error: 'DOCNumber' column not found in input file.")
        return

    # Ensure DOCNumber is properly formatted, handle NaN values
    data['docnumber'] = data['docnumber'].apply(
        lambda x: str(int(float(x))).zfill(8) if pd.notna(x) and x != "" else "MISSING")

    # Get current date for logging purposes
    current_date = datetime.now().strftime('%Y-%m-%d')

    # Ensure required columns are in the DataFrame
    required_columns = ["location", "status", "release date", "date of search"]
    for col in required_columns:
        if col not in data.columns:
            data[col] = "N/A"

    all_subscriber_data = data.copy()  # Create a copy of the original data to maintain consistency

    processed_docs = set()  # Set to track DOCNumbers that have already been processed
    submitted_docs = set()  # Set to track DOCNumbers that have already been submitted for processing

    with ThreadPoolExecutor(max_workers=3) as executor:
        futures = []
        for index, row in data.iterrows():
            if stop_flag():
                logging.info("Stop signal received before scheduling new tasks.")
                break

            doc_number = row['docnumber']

            # Ensure we do not process or submit the same DOCNumber more than once
            if doc_number in processed_docs or doc_number in submitted_docs or doc_number == "MISSING":
                continue

            submitted_docs.add(doc_number)

            first_name = row['firstname']
            last_name = row['lastname']

            # Schedule the process_with_retries function to be run in parallel
            futures.append(
                executor.submit(process_with_retries, index, doc_number, first_name, last_name, stop_flag)
            )

        for future in as_completed(futures):
            if stop_flag():
                logging.info("Stopping all processes as requested.")
                break

            try:
                index, result = future.result()
                doc_number = data.iloc[index]['docnumber']
                if result:
                    # If inmate is found, update the original row data with the new data
                    processed_docs.add(doc_number)
                    all_subscriber_data.at[index, "location"] = result["Location"]
                    all_subscriber_data.at[index, "status"] = result["Status"]
                    all_subscriber_data.at[index, "release date"] = result["Release Date"]
                    all_subscriber_data.at[index, "date of search"] = current_date

                    # Log success and indicate that data was updated
                    logging.info(
                        f"Processed AIC successfully and updated data: DOC number: {result['DOCNumber']}, Name: {result['Name']}")
                    print(f"Processed AIC successfully and updated data: DOC number: {result['DOCNumber']}, Name: {result['Name']}")
                else:
                    # If inmate is not found, retain all original values and update "DATE of SEARCH"
                    processed_docs.add(doc_number)
                    all_subscriber_data.at[index, "date of search"] = current_date

                    # Log not found, indicating no alteration of key data fields
                    logging.info(
                        f"Error finding AIC: DOC number: {doc_number}, Name: {data.iloc[index]['firstname']} {data.iloc[index]['lastname']}. Data retained as original.")
                    print(
                        f"Error finding AIC: DOC number: {doc_number}, Name: {data.iloc[index]['firstname']} {data.iloc[index]['lastname']}. Data retained as original.")
            except Exception as e:
                doc_number = data.iloc[index]['docnumber']
                processed_docs.add(doc_number)
                logging.error(
                    f"Error in processing future result for DOC number {doc_number} ({data.iloc[index]['firstname']} {data.iloc[index]['lastname']}): {e}")

                # Retain original data since processing failed due to an error
                all_subscriber_data.at[index, "date of search"] = current_date

                # Log the error
                logging.info(
                    f"Error processing AIC due to exception: DOC number: {doc_number}, Name: {data.iloc[index]['firstname']} {data.iloc[index]['lastname']}. Data retained as original.")
                print(
                    f"Error processing AIC due to exception: DOC number: {doc_number}, Name: {data.iloc[index]['firstname']} {data.iloc[index]['lastname']}. Data retained as original.")

    # Always output to CSV as requested
    if not all_subscriber_data.empty:
        update_csv(output_file, all_subscriber_data, original_columns=data.columns)
        print(f"Data successfully written to {output_file} with {len(all_subscriber_data)} rows.")
        logging.info(f"Output file created: {os.path.abspath(output_file)}")  # Log full path of the created file
        print(f"Output file created at: {os.path.abspath(output_file)}")  # Explicitly print full path of created file
    else:
        print("No data to write.")

    end_time = time.time()
    elapsed_time = end_time - start_time
    minutes, seconds = divmod(elapsed_time, 60)
    print(f"Time to complete search: {int(minutes)} minutes and {int(seconds)} seconds.")


def update_csv(file_path, data, original_columns):
    if data.empty:
        logging.info("No data to write to CSV.")
        print("No data to write to CSV.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input CSV, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the CSV without an index
    try:
        df.to_csv(file_path, index=False)
        logging.info(f"Data successfully written to {file_path} with {len(data)} rows.")
        print(f"Data successfully written to {file_path} with {len(data)} rows.")
    except Exception as e:
        logging.error(f"Error writing data to CSV: {e}")
        print(f"Error writing data to CSV: {e}")


def update_excel(file_path, data, original_columns):
    print("Excel output is no longer supported. Please use CSV output and convert to Excel if needed.")


def update_csv(file_path, data, original_columns):
    if data.empty:
        logging.info("No data to write to CSV.")
        print("No data to write to CSV.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input CSV, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the CSV without an index
    try:
        df.to_csv(file_path, index=False)
        logging.info(f"Data successfully written to {file_path} with {len(data)} rows.")
        print(f"Data successfully written to {file_path} with {len(data)} rows.")
    except Exception as e:
        logging.error(f"Error writing data to CSV: {e}")
        print(f"Error writing data to CSV: {e}")


def update_excel(file_path, data, original_columns):
    if data.empty:
        logging.info("No data to write to Excel.")
        print("No data to write to Excel.")
        return

    # Create a DataFrame using the original columns to maintain consistency
    df = pd.DataFrame(data)

    # Reorder columns to match the original input Excel, adding missing columns as blank if necessary
    for column in original_columns:
        if column not in df.columns:
            df[column] = ""  # Add missing columns as empty
    df = df[original_columns]

    # Write the DataFrame to an Excel file without index
    try:
        df.to_excel(file_path, index=False, engine='xlsxwriter')
        logging.info(f"Data successfully written to Excel: {file_path}")
        print(f"Data successfully written to Excel: {file_path}")
    except Exception as e:
        logging.error(f"Error writing data to Excel: {e}")
        print(f"Error writing data to Excel: {e}")

