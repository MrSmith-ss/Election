from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from webdriver_manager.firefox import GeckoDriverManager
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import time
import pandas as pd

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By



# Configure Firefox options
options = webdriver.FirefoxOptions()
options.binary_location = r'C:\Program Files\Mozilla Firefox\firefox.exe'  # Change if needed

# Set up Selenium WebDriver for Firefox
service = Service(GeckoDriverManager().install())
driver = webdriver.Firefox(service=service, options=options)

# Dictionary mapping state abbreviations to full names
state_abbr_to_full = {
    "AL": "Alabama", "AK": "Alaska", "AZ": "Arizona", "AR": "Arkansas", "CA": "California",
    "CO": "Colorado", "CT": "Connecticut", "DC": "District of Columbia", "DE": "Delaware", "FL": "Florida", "GA": "Georgia",
    "HI": "Hawaii", "ID": "Idaho", "IL": "Illinois", "IN": "Indiana", "IA": "Iowa",
    "KS": "Kansas", "KY": "Kentucky", "LA": "Louisiana", "ME": "Maine", "MD": "Maryland",
    "MA": "Massachusetts", "MI": "Michigan", "MN": "Minnesota", "MS": "Mississippi", "MO": "Missouri",
    "MT": "Montana", "NE": "Nebraska", "NV": "Nevada", "NH": "New Hampshire", "NJ": "New Jersey",
    "NM": "New Mexico", "NY": "New York", "NC": "North Carolina", "ND": "North Dakota", "OH": "Ohio",
    "OK": "Oklahoma", "OR": "Oregon", "PA": "Pennsylvania", "RI": "Rhode Island", "SC": "South Carolina",
    "SD": "South Dakota", "TN": "Tennessee", "TX": "Texas", "UT": "Utah", "VT": "Vermont",
    "VA": "Virginia", "WA": "Washington", "WV": "West Virginia", "WI": "Wisconsin", "WY": "Wyoming"
}

# Function to scrape party and vote data for a given state URL
def scrape_state_data(url, state_abbreviation):
    driver.get(url)

    # Wait until the iframe is available
    iframe = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, "//iframe[@frameborder='0' and @loading='lazy' and @width='100%']"))
    )
    # Switch to the iframe
    driver.switch_to.frame(iframe)
    
    # Scrape the content inside the iframe
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    # Extract party and vote information
    parties = soup.find_all("span", class_="rounded-sm bg-neutral-200 px-2 py-1 text-xs font-semibold text-stone-800")
    votes = soup.find_all("td", class_="pr-1 text-end text-base sm:pr-2")

    # Map party names to abbreviations and create the results list
    results = []
    for party, vote in zip(parties, votes):
        party_name = party.text.strip()
        if party_name == "GOP":
            party_abbr = "R"
        elif party_name == "DEM":
            party_abbr = "D"
        else:
            party_abbr = "I"  # For any other party

        results.append((state_abbreviation, party_abbr, vote.text.strip()))

    return results

# Prepare a list to store the data for all states
all_results = []

# Loop through each state in the dictionary
for state_abbr, state_full in state_abbr_to_full.items():
    # Generate the URL dynamically, lowercasing the full state name and replacing spaces with dashes
    state_url = f"https://decisiondeskhq.com/results/2024/General/races/{state_full.lower().replace(' ', '-')}-president-all-parties-general-election"
    
    print(f"Scraping results for {state_full} ({state_abbr})...")
    state_results = scrape_state_data(state_url, state_abbr)
    all_results.extend(state_results)

# Close the driver
driver.quit()

# Create a DataFrame from the collected data
df = pd.DataFrame(all_results, columns=["STATE ABBREVIATION", "PARTY", "GENERAL RESULTS"])

# Export the DataFrame to an Excel file
df.to_excel("2024_Election_Results.xlsx", index=False)


