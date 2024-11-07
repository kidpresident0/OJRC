from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import csv

# Function to run the main web scraping process
def run_main_process(input_file, output_file, stop_flag_callback):
    # Set up the web driver (using Chrome in this example)
    driver = webdriver.Chrome()  # Make sure you have the ChromeDriver installed and in PATH
    driver.maximize_window()
    wait = WebDriverWait(driver, 10)

    try:
        # Navigate to the login page
        driver.get("https://visor.oregon.gov/")

        # Input Username
        username_field = wait.until(EC.presence_of_element_located((By.ID, "UserName")))
        username_field.send_keys("doitforbooterz@gmail.com")

        # Input Password
        password_field = driver.find_element(By.ID, "Password")
        password_field.send_keys("10CharacterOJRC!")

        # Submit the form by pressing enter (or you could click the login button)
        password_field.send_keys(Keys.RETURN)

        # Wait for login to complete and search box to appear
        wait.until(EC.presence_of_element_located((By.ID, "searchText")))

        # Read the DOC numbers from the input CSV file
        with open(input_file, "r") as csvfile:
            reader = csv.reader(csvfile)
            headers = next(reader)  # Skip header row if there's one
            doc_numbers = [row[0] for row in reader]

        # Prepare output CSV file
        with open(output_file, "w", newline="") as outfile:
            writer = csv.writer(outfile)
            writer.writerow(["DOC Number", "Status", "Location", "Release Date"])

            # Iterate over DOC numbers
            for doc_number in doc_numbers:
                if stop_flag_callback():
                    print("Stop flag detected, ending the process.")
                    break

                # Input the DOC number in the search box
                search_box = driver.find_element(By.ID, "searchText")
                search_box.clear()
                search_box.send_keys(doc_number)

                # Click the search button
                search_button = driver.find_element(By.ID, "searchBtn")
                search_button.click()

                try:
                    # Wait for the search results to load
                    details_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(@class, 'k-grid-Details')]")))
                    details_button.click()

                    # Extract inmate details (status, location, release date)
                    status = wait.until(EC.presence_of_element_located((By.ID, "idoc_offenderstatus_readonly"))).text
                    location = driver.find_element(By.ID, "idoc_name_readonly").text
                    release_date = driver.find_element(By.ID, "idoc_eprd_readonly").text

                    # Write the details to the output CSV
                    writer.writerow([doc_number, status, location, release_date])
                    print(f"Processed DOC Number: {doc_number}")

                except TimeoutException:
                    print(f"Details not found for DOC Number: {doc_number}")

                # Pause briefly to avoid overwhelming the server
                time.sleep(2)

    except Exception as e:
        print(f"An error occurred: {e}")

    finally:
        # Close the browser
        driver.quit()
