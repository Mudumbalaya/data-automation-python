import requests
from bs4 import BeautifulSoup
import re
import os
import pandas as pd

print("🚀 Script started")

file_path = "C:\\Users\\Layacherry\\OneDrive\\Desktop\\founders\\companies(1).xlsx.xlsx"

print("🔍 Checking for Excel file at:", file_path)

if not os.path.exists(file_path):
    print(f"❌ File not found at:\n{file_path}")
    exit()

try:
    df = pd.read_excel(file_path)
    print("📊 Total rows in Excel file:", len(df))
    print(df.head())
except Exception as e:
    print(f"❌ Error reading the Excel file: {str(e)}")
    exit()

if 'Company URL' not in df.columns:
    print("❌ 'Company URL' column not found.")
    exit()

# Helper function to detect Contact Us pages based on common paths
def find_contact_page(url):
    contact_paths = ['/contact', '/contact-us', '/about']
    for path in contact_paths:
        contact_url = url.rstrip('/') + path
        try:
            response = requests.get(contact_url, timeout=5)
            if response.status_code == 200:
                print(f"📍 Found contact page: {contact_url}")
                return contact_url
        except requests.RequestException:
            continue
    return url  # If no contact page found, return the original URL

# Improved address extraction using patterns for common components
def extract_addresses(text):
    address_patterns = [
        r'\b(?:\d{5}|\d{6})\b',  # PIN codes
        r'\b(?:India|Street|City|State|Road)\b',  # Location-related keywords
        r'\b(?:[A-Za-z]+(?: [A-Za-z]+)* [A-Za-z]+(?: [A-Za-z]+)*)\b',  # Generic address patterns
    ]
    
    addresses = []
    for pattern in address_patterns:
        addresses.extend(re.findall(pattern, text, re.IGNORECASE))
    
    return addresses

def extract_contact_info(url):
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                      "AppleWebKit/537.36 (KHTML, like Gecko) "
                      "Chrome/90.0.4430.93 Safari/537.36"
    }

    try:
        # First, try to find the Contact Us page if not already given
        contact_url = find_contact_page(url)

        response = requests.get(contact_url, headers=headers, timeout=10)
        print(f"🌐 {contact_url} → Status: {response.status_code}")

        if response.status_code != 200:
            return {"error": f"HTTP {response.status_code}"}

        soup = BeautifulSoup(response.text, 'html.parser')
        text = soup.get_text()
        print(f"📄 Text length: {len(text)} characters")

        # Extracting emails, phones, and addresses
        emails = re.findall(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", text)
        phones = re.findall(r"\+?\d[\d\s().-]{8,}\d", text)
        addresses = extract_addresses(text)

        return {
            "emails": list(set(emails)),
            "phones": list(set(phones)),
            "addresses": list(set(addresses))
        }

    except Exception as e:
        return {"error": str(e)}

contact_info_list = []

for index, row in df.iterrows():
    company_url = row['Company URL']
    if pd.isna(company_url):
        print(f"⚠️ Skipping row {index + 1} due to missing URL.")
        continue

    print(f"\n🔍 Extracting from: {company_url}")

    contact_info = extract_contact_info(company_url)

    contact_info_list.append({
        "Company Name": row['Company Name'],
        "URL": company_url,
        "Emails": ', '.join(contact_info['emails']) if 'emails' in contact_info else "❌ No emails found",
        "Phones": ', '.join(contact_info['phones']) if 'phones' in contact_info else "❌ No phones found",
        "Addresses": ', '.join(contact_info['addresses']) if 'addresses' in contact_info else "❌ No addresses found",
        "Error": contact_info.get("error", "")
    })

# Save to Excel
output_file = "C:\\Users\\Layacherry\\OneDrive\\Desktop\\founders\\enhanced_extracted_contact_info.xlsx"
print(f"\n💾 Saving data to: {output_file}")
pd.DataFrame(contact_info_list).to_excel(output_file, index=False)
print("✅ Extraction complete.")
