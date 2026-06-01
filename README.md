📋 Task Manager Web App
> A role-based task management system built with Google Apps Script — featuring separate Director and Employee portals, automated email notifications, and a performance scoring dashboard.
![Google Apps Script](https://img.shields.io/badge/Google%20Apps%20Script-4285F4?style=for-the-badge&logo=google&logoColor=white)
![JavaScript](https://img.shields.io/badge/JavaScript-F7DF1E?style=for-the-badge&logo=javascript&logoColor=black)
![HTML](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![Google Sheets](https://img.shields.io/badge/Google%20Sheets-34A853?style=for-the-badge&logo=google-sheets&logoColor=white)
---
📌 Overview
This web application was built to streamline task delegation and tracking within a team. It provides two separate role-based portals — one for Directors (managers) and one for Employees — each with their own login credentials and tailored interface.
Directors can assign tasks, set deadlines, and monitor team performance through a scoring dashboard. Employees receive instant email notifications when tasks are assigned and can update their task status in real time — all backed by Google Sheets as a live database, with no external hosting or infrastructure required.
---
✨ Features
🔐 Role-based login system — separate secure portals for Directors and Employees
📝 Task assignment — Directors can create, assign, and track tasks with deadlines
📧 Automated email notifications — employees are notified instantly on task assignment
📊 Performance dashboard — scoring system to track employee task completion rates
📱 Mobile-friendly design — works on phones, tablets, and desktops
📖 User manuals included — Excel-based reference guides for both Director and Employee roles
---
📁 File Structure
```
Task-Manager-Webapp/
├── Code_Github.gs           # Google Apps Script backend (auth, task logic, email, data)
├── director.html            # Director portal UI (task assignment, dashboard, team view)
├── employee.html            # Employee portal UI (task list, status updates)
├── Director_Manual.xlsx     # User manual for Directors
├── Employee_Manual.xlsx     # User manual for Employees
└── README.md
```
---
⚙️ How It Works
Director Flow
Director logs in with their credentials
Views the team dashboard with task status and performance scores
Assigns new tasks to employees — sets title, description, priority, and deadline
Employee receives an automatic email notification
Director tracks completion in real time via the dashboard
Employee Flow
Employee logs in with their unique ID and password
Views their assigned task list with priorities and deadlines
Updates task status as work progresses
Completed tasks are reflected immediately in the Director's dashboard
---
🚀 Setup & Deployment
Prerequisites
A Google account with access to Google Sheets and Gmail
Basic familiarity with Google Apps Script (script.google.com)
Steps
Create a new Google Spreadsheet
Set up sheets for: Users, Tasks, and Performance data (refer to the included manuals)
Open Apps Script
In the spreadsheet, go to Extensions → Apps Script
Copy the code files
Paste `Code_Github.gs` into the script editor
Create two HTML files named `director` and `employee`, and paste the respective HTML content
Update the configuration in `Code_Github.gs`:
```javascript
   const SPREADSHEET_ID = 'YOUR_GOOGLE_SHEET_ID';
   ```
Deploy as a Web App
Click Deploy → New Deployment
Type: Web App
Execute as: Me
Who has access: Anyone within your organisation (or as required)
Click Deploy and share the URL with your team
Refer to the manuals
`Director_Manual.xlsx` — covers how to set up user accounts and use the Director portal
`Employee_Manual.xlsx` — covers how employees log in and manage their tasks
---
🛠️ Tech Stack
Layer	Technology
Frontend	HTML5, CSS3, JavaScript
Backend	Google Apps Script (V8 runtime)
Database	Google Sheets
Notifications	Gmail via Apps Script `MailApp`
Distribution	Google Apps Script Web App deployment
---
📖 Documentation
User manuals are included in the repository for both roles:
📘 Director Manual — Setup, task assignment, dashboard usage
📗 Employee Manual — Login, viewing tasks, updating status
---
👤 Author
Avinaba Chakraborty
Senior MIS Analyst · 16+ Years Experience in BI & Data Analytics
LinkedIn · GitHub
---
📄 License
This project is open source and available under the MIT License.
