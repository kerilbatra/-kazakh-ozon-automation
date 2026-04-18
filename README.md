# Kazakh Ozon Automation

This project automates the process of fetching financial realization data from the Ozon Seller API, processing it, and syncing it with a SharePoint Excel file.

---

## 🚀 Features

- Fetches monthly financial realization report from Ozon API
- Automatically calculates previous month data
- Flattens and structures nested JSON response
- Converts data into a clean pandas DataFrame
- Removes duplicate records
- Appends only new records to existing Excel file
- Uploads updated file back to SharePoint
- Automatically cleans up local temporary files

---

## 🛠️ Tech Stack

- Python
- Requests (API calls)
- Pandas (data processing)
- OpenPyXL (Excel handling)
- Office365-REST-Python-Client (SharePoint integration)

---

## 📁 Project Workflow

1. Connects to Ozon API  
2. Retrieves financial report for previous month  
3. Processes and cleans data  
4. Downloads existing SharePoint Excel file  
5. Merges new + existing data  
6. Removes duplicates  
7. Updates Excel file  
8. Uploads back to SharePoint  
9. Deletes local temporary file  
