import streamlit as st
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
import random
import os
import time

# ----------------------------
# PAGE TITLE
# ----------------------------
st.set_page_config(page_title="Lead Scrapper Tool")

st.title("Lead Scrapper Tool")

# ----------------------------
# SESSION STATE
# ----------------------------
if "stop_scraping" not in st.session_state:
    st.session_state.stop_scraping = False

# ----------------------------
# URL INPUT
# ----------------------------
url = st.text_input(
    "Enter URL",
    placeholder="https://www.justdial.com/Chennai/Clinics/nct-10101647"
)

# ----------------------------
# BUTTONS
# ----------------------------
col1, col2 = st.columns(2)

with col1:
    start_button = st.button("Start Scraping")

with col2:
    stop_button = st.button("Stop Scraping")

# ----------------------------
# STOP BUTTON
# ----------------------------
if stop_button:
    st.session_state.stop_scraping = True
    st.warning("Scraping stopped by user")
    st.stop()

# ----------------------------
# START BUTTON
# ----------------------------
if start_button:
    st.session_state.stop_scraping = False

if not start_button:
    st.stop()

# ----------------------------
# VALIDATE URL
# ----------------------------
if not url:
    st.error("Please enter URL")
    st.stop()

# ----------------------------
# START CHROME
# ----------------------------
driver = uc.Chrome()

processed_urls = set()

try:

    driver.get(url)

    st.write("Page Title :", driver.title)
    st.write("Current URL :", driver.current_url)

    time.sleep(5)

    # ----------------------------
    # SCROLL PAGE
    # ----------------------------
    last_height = driver.execute_script(
        "return document.body.scrollHeight"
    )

    for i in range(5):

        if st.session_state.stop_scraping:
            driver.quit()
            st.warning("Scraping stopped")
            st.stop()

        st.write(f"Scrolling : {i + 1}")

        driver.execute_script(
            "window.scrollTo(0, document.body.scrollHeight);"
        )

        time.sleep(random.randint(5, 10))

        new_height = driver.execute_script(
            "return document.body.scrollHeight"
        )

        if new_height == last_height:
            st.write("All listings loaded")
            break

        last_height = new_height

    # ----------------------------
    # GET PAGE SOURCE
    # ----------------------------
    soup = BeautifulSoup(driver.page_source, "html.parser")

    # ----------------------------
    # FIND BUSINESS LINKS
    # ----------------------------
    business_links = []

    all_links = soup.find_all("a", href=True)

    for link in all_links:

        href = link["href"]

        if "/Chennai/" in href and "nct-" not in href:

            full_url = href

            if full_url.startswith("/"):
                full_url = "https://www.justdial.com" + full_url

            if full_url not in business_links:
                business_links.append(full_url)

    st.write("Total business links :", len(business_links))

    if len(business_links) == 0:
        st.error("No business links found")
        driver.quit()
        st.stop()

    # ----------------------------
    # EXCEL FILE
    # ----------------------------
    excel_file = "justdial_leads.xlsx"

    if os.path.exists(excel_file):

        wb = load_workbook(excel_file)
        ws = wb.active

        st.write("Existing Excel file opened")

    else:

        wb = Workbook()
        ws = wb.active

        ws.append([
            "Business Name",
            "Website",
            "Justdial URL",
            "Phone Number",
            "Address"
        ])

        wb.save(excel_file)

        st.write("New Excel file created")

    st.write("Excel Path :", os.path.abspath(excel_file))

    # ----------------------------
    # PROGRESS BAR
    # ----------------------------
    progress_bar = st.progress(0)

    status_text = st.empty()

    # ----------------------------
    # OPEN BUSINESS PAGES
    # ----------------------------
    count = 0

    for business_url in business_links:

        # ----------------------------
        # STOP SCRAPING
        # ----------------------------
        if st.session_state.stop_scraping:
            driver.quit()
            st.warning("Scraping stopped")
            st.stop()

        count += 1

        # ----------------------------
        # UPDATE PROGRESS
        # ----------------------------
        progress = count / len(business_links)

        progress_bar.progress(progress)

        status_text.write(
            f"Processing {count} / {len(business_links)}"
        )

        # ----------------------------
        # SKIP DUPLICATES
        # ----------------------------
        if business_url in processed_urls:
            continue

        # ----------------------------
        # RESTART DRIVER
        # ----------------------------
        if count % 20 == 0:

            st.write("Restarting browser...")

            driver.quit()

            time.sleep(5)

            driver = uc.Chrome()

            time.sleep(5)

        try:

            driver.get(business_url)

            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.TAG_NAME, "title")
                )
            )

            detail_soup = BeautifulSoup(
                driver.page_source,
                "html.parser"
            )

            # ----------------------------
            # BUSINESS NAME
            # ----------------------------
            business_name = (
                driver.title.split("in")[0].strip()
            )

            # ----------------------------
            # ALL LINKS
            # ----------------------------
            detail_links = detail_soup.find_all(
                "a",
                href=True
            )

            # ----------------------------
            # PHONE NUMBER
            # ----------------------------
            phone_number = "Not Found"

            for dlink in detail_links:

                text = dlink.text.strip()

                if text.isdigit() and len(text) >= 10:

                    phone_number = text
                    break

            # ----------------------------
            # ADDRESS
            # ----------------------------
            address = "No Address"

            address_tag = detail_soup.find(
                "a",
                class_="color111"
            )

            if address_tag:

                address = address_tag.get_text(
                    " ",
                    strip=True
                )

            # ----------------------------
            # WEBSITE
            # ----------------------------
            website = "No Website"

            for dlink in detail_links:

                dhref = dlink["href"].lower()

                if (
                    "http" in dhref
                    and "justdial" not in dhref
                    and "facebook" not in dhref
                    and "instagram" not in dhref
                    and "youtube" not in dhref
                    and "whatsapp" not in dhref
                ):

                    website = dhref
                    break

            # ----------------------------
            # CLEAN URL
            # ----------------------------
            clean_url = business_url.split("?")[0]

            if clean_url in processed_urls:
                continue

            # ----------------------------
            # SAVE TO EXCEL
            # ----------------------------
            ws.append([
                business_name,
                website,
                clean_url,
                phone_number,
                address,
            ])

            wb.save(excel_file)

            processed_urls.add(clean_url)

        except Exception as e:

            st.error(f"Error : {e}")

    # ----------------------------
    # CLOSE DRIVER
    # ----------------------------
    driver.quit()

    # ----------------------------
    # SUCCESS
    # ----------------------------
    st.success("Scraping Completed")

    st.write("Excel saved :", excel_file)

    # ----------------------------
    # DOWNLOAD BUTTON
    # ----------------------------
    if os.path.exists(excel_file):

        with open(excel_file, "rb") as file:

            st.download_button(
                label="Download Excel",
                data=file,
                file_name=excel_file,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

except Exception as e:

    st.error(f"Main Error : {e}")

    try:
        driver.quit()
    except:
        pass 

# import streamlit as st

# st.title("Working")