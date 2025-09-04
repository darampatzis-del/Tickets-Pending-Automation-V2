### Overview
This Python script automates the process of **Tickets Pending** Daily Task.
### Functionality
1.  Reads ticket data from an OTRS generated Excel source file.
2.  Uses a predefined Excel template to generate a new output excel file with the transformed data.
3.  Detects and tags high-priority tickets by scanning ticket subjects for specific keywords.
4.  Sorts tickets by priority and creation date.
5.  Categorizes Tickets into Sheets and automatically assigns tickets to different sheets based on the “Queue” field.
### Output
New excel file **“Tickets Pending \<current-date\>.xlsx“** that contains:
-  Sheet **“All”** with all tickets formatted and sorted by Priority and Date Created.
-  Categorized sheets (**SD_CS, MM_PP_QM, FI-CO, System**) with relevant tickets based on **Queue** column.
## Method 2 - WSL (Contains errors | Not recommended)
### Prerequisites
1.  Make sure that **Windows Subsystem for Linux (WSL)** is enabled.
![Tickets Pending Guide](https://github.com/user-attachments/assets/efdf29e6-8043-47d3-aa99-29b6e1d3dde9)
2.  Files **Template.xlsx, Queues.txt and Customers.txt** must be in the same folder as the script.
### Installation
1.  Ensure that **Windows Subsystem for Linux (WSL) is enabled**.
2.  Install **Ubuntu** from **Microsoft Store**
3.  In the Ubuntu execute the following commands:
    -  **sudo apt update -y**
    -  **sudo apt install python3-pandas python3-openpyxl -y**
### Execution Instructions
1.  Launch Ubuntu instance.
2.  Navigate to the location where the script is placed with the following command:
    +  **cd “/mnt/c/PATH_TO_LOCATION”** - for example if script is placed to Desktop/Tickets Pending Automation, the path should be:
    +  **cd “/mnt/c/Users/<username>/Desktop/Tickets Pending Automation”**
3.  Make sure that you download latest excel file from OTRS and place it to the same folder as the script.
4.  Execute script with the following command:
    +  **python3 tp_automation.py <generated_OTRS_file.xlsx>**\
**Note: If you face any issues executing the script, try to execute as root user with the below command**
    +  **sudo python3 tp_automation.py <generated_OTRS_file.xlsx>**
5.  After executing script, a new file named “Tickets Pending <current_date>.xlsx” will be created. The file contains errors (working to fix it), select yes to let excel repair it and when it opens, save it again to have a non-error file.\
**Note: If you see any empty values in Customer or Queue column, alter the file Customers.txt or Queues.txt as there you can add all the customers and queues.**
### Notes
Please let me know if you need further information and feel free to enhance the script and fix any issues that may arise.
