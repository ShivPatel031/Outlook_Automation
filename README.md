# 📧 Outlook Automation with Python

This project automates the process of fetching, categorizing, and organizing Outlook emails using the **Microsoft Graph API**. It supports handling attachments, managing mail folders, and applying categories — ideal for teams dealing with alert-based or operational email workflows.

---

## 📂 Project Structure

| File              | Purpose                                                       |
|-------------------|---------------------------------------------------------------|
| `ms_graph.py`     | Handles authentication and connection to Microsoft Graph API. |
| `outlook.py`      | Contains reusable functions to interact with Outlook API.     |
| `pre_alert.py`    | The **main automation script** that processes and sorts emails. |
| Other files       | Practice or test files (can be ignored).                      |

---

## ⚙️ How `pre_alert.py` Works

The `pre_alert.py` script is the **core automation engine** that performs the entire workflow — from token generation to downloading attachments and moving emails into specific folders.

### 🔄 Step-by-Step Workflow

1. **Token Initialization**
   - Uses `ms_graph.py` to authenticate and retrieve an access token from Microsoft Graph API.

2. **Check Last Processed Time**
   - Checks for a `last_outlook_check_time.txt` file:
     - If found: reads the time and adds 1 second to avoid overlap.
     - If not found: defaults to **2 days ago** from the current time.
   - This ensures only **new emails** are fetched.

3. **Fetch and Filter Emails**
   - Retrieves emails from the inbox after the defined time.
   - Filters only those matching a **predefined Pre-Alert criteria**.

4. **Prepare Mail Folders**
   - Ensures the following folders exist in Outlook:
     - `PreAlert`
     - `No Attachment`
     - `Query`
   - If missing, the script **creates them automatically**.

5. **Process Emails**

   For each filtered email:
   - 📩 **From specific sender** (`210303105085@paruluniversity.ac.in`):
     - Add **Yellow** category.
     - Move to **Query** folder.
   - 📎 **With attachment**:
     - Download attachments to:
       ```
       ./downloaded/<subject>_<received_time>/
       ```
     - Add **Orange** category.
     - Move to **PreAlert** folder.
   - ❗ **Without attachment** (but marked as Pre-Alert):
     - Add **Orange** and **Yellow** categories.
     - Move to **No Attachment** folder.

6. **Update Last Checked Time**
   - Writes the latest processed email's timestamp to `last_outlook_check_time.txt`.

---

## 🗂 Folder Logic Summary

| Folder Name      | When Used                                | Assigned Categories    |
|------------------|-------------------------------------------|-------------------------|
| `PreAlert`       | Email has valid attachments               | 🟠 Orange               |
| `No Attachment`  | Email is Pre-Alert but missing attachments| 🟠 Orange, 🟡 Yellow     |
| `Query`          | Needs human intervention or specific user | 🟡 Yellow               |

---

## 🚀 Features

- ✅ Automatically fetches new emails using time-based filtering
- ✅ Detects and downloads attachments to organized local folders
- ✅ Assigns Outlook categories for easy filtering
- ✅ Moves emails into relevant Outlook folders
- ✅ Auto-creates folders if they don’t exist
- ✅ Fully integrated with **Microsoft Graph API**

---

## 🔐 Prerequisites

- Python 3.7+
- Microsoft 365 account with API access
- Microsoft Graph API permissions (Mail.ReadWrite, Mail.Send, etc.)

---

## 📦 Setup Instructions

1. **Clone the repository**
   ```bash
   git clone https://github.com/ShivPatel031/Outlook_Automation.git
   cd Outlook_Automation

🧠 Future Improvements

✅ Add multithreading to speed up processing (planned)

📁 Add support for other folder types (optional)

📊 Dashboard/summary of processed emails

