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
# CATEGORY PAGE
# ----------------------------
st.title("lead Scrapper Tool")
url=st.text_input("Enter  url ",placeholder="https://www.justdial.com/Chennai/Clinics/nct-10101647")
start_button=st.button("start Scrapping")
if not start_button:
    st.stop()
if not url:
    st.error("enter url ")
    st.stop()
url="https://www.justdial.com/Chennai/Clinics/nct-10101647"

driver = uc.Chrome()

processed_urls = set()

driver.get(url)

print(driver.title)
print("URL :", driver.current_url)

time.sleep(5)

# ----------------------------
# SCROLL PAGE
# ----------------------------
last_height = driver.execute_script(
    "return document.body.scrollHeight"
)

for i in range(5):

    st.write("Scrolling :", i + 1)

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
# GET PAGE HTML
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

# ----------------------------
# CREATE / OPEN EXCEL FILE
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
# OPEN EACH BUSINESS PAGE
# ----------------------------
count = 0

for business_url in business_links:

    count += 1

    if business_url in processed_urls:
        continue

    # Restart browser every 20 pages
    if count % 20 == 0:

        st.write("Restarting browser...")

        driver.quit()

        time.sleep(5)

        driver = uc.Chrome()

        time.sleep(5)

    try:

        st.write("\nOpening :", business_url)

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
        business_name = driver.title.split("in")[0].strip()

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

        # st.write("Saved :", website)

    except Exception as e:

        st.write("Error :", e)

# ----------------------------
# CLOSE DRIVER
# ----------------------------
with open(excel_file, "rb") as file:

    st.download_button(
        label="Download Excel",
        data=file,
        file_name=excel_file,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
driver.quit()

st.success("\nCompleted")
st.write("Excel saved :", excel_file)