# 🚀 Overview

This project is a fully automated Employee Shift Management System built in Microsoft Excel, designed to solve real-world scheduling challenges for small to medium-sized businesses.

It simulates a real business scenario of a restaurant owner managing multiple staff roles, working hours, wage rules, and fairness in shift allocation.

The solution transforms Excel from a simple spreadsheet into a decision-support system using advanced features such as formulas, Power Query, Power Pivot (DAX), and VBA. The solution also works with older version of Excel making it appealing to a wider Excel enthusiasts.

# 🎯 Business Problem

Managing employee shifts manually often leads to:

❌ Overworked staff exceeding legal hour limits

❌ Unfair distribution of weekend and holiday shifts

❌ Payroll errors due to complex wage rules

❌ Lack of visibility into staffing coverage

❌ Scheduling conflicts like “CLOPEN” (Late night followed by early morning)

This project addresses all of the above with an automated and scalable solution.

# 💡 Solution Highlights
## 🧠 Smart Scheduling Engine
* Dynamic staff selection with automated data retrieval
* Cascading dropdowns based on role-specific shifts
* Prevents invalid shift assignments
## ⏱️ Working Hours Control
* Automatic calculation of daily and weekly working hours
* Alerts when staff exceed maximum allowed hours
* CLOPEN shift detection for employee well-being
## 💰 Automated Wage Calculation
* Configurable pay rules:
  * Weekday (normal rate)
  * Weekend (1.5x multiplier)
  * Public Holiday (2x multiplier)
* Fully automated weekly wage computation
## ⚖️ Fairness & Workforce Analytics
* Tracks premium shifts (weekend/holiday) per employee
* Identifies unfair workload distribution
* Uses historical data for trend analysis
## 📊 Interactive Dashboard
* KPI metrics (headcount, wage, average rate)
* Workforce distribution by role
* Workload vs contract hours visualization
* Premium shift distribution insights
* Fully interactive with slicers and timeline

# 🛠️ Technical Implementation
## Excel Features
* Advanced formulas: INDEX, MATCH, IF, COUNTIFS, WEEKDAY
* Dynamic Named Ranges
* Conditional Formatting (rule-based alerts)
## Data Transformation
* Power Query
  * Data cleaning and transformation
  * Unpivoting cross-tab data into analytical format
## Data Modeling
* Power Pivot (DAX) as alternative solution to Dynamic Array Formula in newer Excel version
* Custom measures (e.g., premium shift tracking) for effective visualization
* Aggregations and business logics
## VBA
* Data archiving (weekly → historical)
* Dashboard behavior enhancements
* Chart axis correction to ensure effective visualization

# 📁 Project Structure
* **Sheet Ref**: Reference tables (Roles, Shift Types)
* **Sheet Staff**: Employee master data
* **Sheet Schedule**: Weekly scheduling engine
* **Sheet PublicHoliday**: Holiday reference
* **Sheet Setting**: Configurable parameters (rates, inputs)
* **Sheet tblWeeklyData**: Data transformation from Sheet Schedule using Power Query into big table for dashboard and visualization
* **Sheet tblHistory**: Archived historical data (via VBA)
* **Sheet Pivot**: Power Pivot Tables for all dashboard graphs and charts
* **Sheet Dashboard**: Interactive reporting interface

# 📈 Key Features Demonstrated
* End-to-end system design in Excel
* Data modeling and transformation
* Business rule implementation
* Automation and user experience design
* Analytical dashboard creation

# 🎥 Demo / Walkthrough
(Optional – Add your YouTube video link here)

# 🧑‍💼 Why This Project Matters

This project demonstrates the ability to:

* Translate real business problems into technical solutions
* Design scalable and maintainable data models
* Combine multiple Excel tools into one cohesive system
* Deliver insights, not just data with interactive dashboard
* Customize to fit with other business models such as **Nurse Shift** and **Warehouse Shift Management**

# 📌 Potential Enhancements
* Integration with external data sources (SQL / APIs)
* Migration to Power BI for enterprise-level reporting
* Optimization using modern Excel functions (e.g., dynamic arrays)

# 🤝 Contact

If you’re interested in discussing this project or opportunities:

* LinkedIn: https://www.linkedin.com/in/vengm/
* Email: khmer.keang@gmail.com
