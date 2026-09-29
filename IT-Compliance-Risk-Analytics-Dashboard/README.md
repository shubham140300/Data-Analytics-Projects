# IT Compliance & Risk Analytics Dashboard

An end-to-end IT infrastructure compliance and risk analytics project designed to analyze server compliance data, identify potentially high-risk servers, and present actionable insights through interactive Power BI and Streamlit dashboards.

---

## 1. Project Overview

IT infrastructure environments generate large volumes of compliance, monitoring, and security data from multiple enterprise tools.

This project demonstrates an end-to-end analytics workflow for transforming infrastructure compliance data into meaningful insights.

The project combines:

- SQL
- MySQL
- Excel
- Python
- Pandas
- Power BI
- DAX
- Plotly
- Streamlit
- GitHub
- Streamlit Community Cloud

The solution focuses on analyzing infrastructure compliance, identifying potentially high-risk servers, understanding risk patterns, and presenting the results through interactive dashboards.

The project contains two dashboarding layers:

1. Power BI Dashboard
2. Streamlit Web Dashboard

---

## 2. Business Problem

Large enterprise IT environments can contain thousands of servers running different operating systems and supporting different applications.

These servers may be monitored by multiple security, compliance, and infrastructure tools such as:

- Splunk
- Qualys
- CrowdStrike
- ServiceNow
- TGIM
- RSA
- Logger

Each tool can provide different information about the infrastructure.

Analyzing these sources manually can make it difficult to quickly answer questions such as:

- Which servers are non-compliant?
- Which servers are potentially high-risk?
- Which priority contains the highest number of high-risk servers?
- Which operating system has more high-risk servers?
- Which environment contains more high-risk servers?
- Which incident categories are associated with high-risk servers?
- What percentage of analyzed servers are classified as high-risk?
- How can the information be presented in an interactive and easy-to-understand format?

This project addresses these requirements by transforming infrastructure compliance data into an analytical dataset and presenting the results through interactive dashboards.

---

## 3. Project Objective

The main objectives of the project are:

- Analyze infrastructure compliance data.
- Clean and transform raw infrastructure data.
- Combine information from different compliance and security sources.
- Classify infrastructure records based on risk.
- Identify potentially high-risk servers.
- Calculate high-risk server counts.
- Calculate high-risk rates.
- Analyze risk by priority.
- Analyze risk by operating system.
- Analyze risk by environment.
- Analyze risk by incident category.
- Build an interactive Power BI dashboard.
- Build an interactive Streamlit dashboard.
- Deploy the Streamlit dashboard using Streamlit Community Cloud.
- Create a portfolio project demonstrating practical data analytics and business intelligence skills.

---

## 4. Why This Project Matters

Infrastructure teams need to identify compliance and security issues efficiently.

A large raw dataset can be difficult to interpret manually.

An analytical dashboard provides a faster way to understand:

- Total analyzed servers
- High-risk server count
- High-risk percentage
- Priority-wise risk
- Operating-system-wise risk
- Environment-wise risk
- Incident-category-wise risk

This allows analysts and infrastructure teams to investigate areas with higher concentrations of risk and support remediation activities.

---

## 5. End-to-End Architecture

The project follows the following overall workflow:

```text
Raw Infrastructure Compliance Data
                |
                v
          Data Cleaning
                |
                v
       Data Transformation
                |
                v
        Risk Classification
                |
                v
       Prediction / Risk Output
                |
        +-------+-------+
        |               |
        v               v
    Power BI        Streamlit
    Dashboard       Dashboard
        |               |
        v               v
 Business Insights   Interactive
                     Web Analytics
```

---

## 6. Project Workflow

The complete project workflow consists of the following stages.

### Step 1 - Data Collection

Infrastructure compliance information is collected from multiple monitoring, security, and compliance sources.

### Step 2 - Data Cleaning

The raw dataset is checked for:

- Missing values
- Duplicate records
- Incorrect data types
- Inconsistent naming
- Invalid values
- Data-quality issues

### Step 3 - Data Transformation

The raw data is transformed into a structured analytical dataset.

Transformation activities include:

- SQL `CASE WHEN`
- SQL `UNION ALL`
- Data type conversion
- Standardization
- Filtering
- Aggregation
- Risk classification

### Step 4 - Risk Classification

Infrastructure records are classified based on compliance and security-related attributes.

The resulting analytical dataset contains a `predicted_high_risk` field.

### Step 5 - Data Analysis

The transformed dataset is analyzed using SQL, Python, and Power BI.

### Step 6 - Dashboard Development

Interactive dashboards are developed using Power BI and Streamlit.

### Step 7 - Deployment

The Streamlit dashboard is deployed using Streamlit Community Cloud and connected to the GitHub repository.

---

## 7. Source Data

The project uses infrastructure compliance information containing server-level attributes and compliance/security indicators.

The raw dataset contains fields such as:

```text
Server_ID
Hostname
IP_Address
OS_Type
Assignment_Team
Environment
Scan_Date
Windows_Hardening
CrowdStrike
SNOW_Status
Splunk_Status
TGIM
RSA_Status
Logger
Qualys
```

These fields represent different infrastructure and compliance characteristics.

### Server_ID

Unique identifier of the server.

### Hostname

Name assigned to the server.

### IP_Address

Network address associated with the server.

### OS_Type

Operating system installed on the server.

Examples can include:

- Windows
- Red Hat Enterprise Linux

### Assignment_Team

Team responsible for the server.

### Environment

Environment classification such as:

- Production
- Development
- Test

### Scan_Date

Date on which the compliance or security information was collected.

### Windows_Hardening

Represents the Windows hardening compliance status where applicable.

### CrowdStrike

Represents endpoint security status.

### SNOW_Status

Represents ServiceNow-related status information.

### Splunk_Status

Represents Splunk monitoring/logging status.

### TGIM

Represents the corresponding infrastructure compliance/security check.

### RSA_Status

Represents RSA-related security information.

### Logger

Represents logging-related status.

### Qualys

Represents vulnerability/compliance information from Qualys.

---

## 8. Data Cleaning

Before performing analytics, the raw data needs to be cleaned and standardized.

Typical data-cleaning activities include:

- Removing duplicate records
- Handling missing values
- Standardizing column names
- Standardizing categorical values
- Converting data types
- Validating server identifiers
- Validating environment values
- Validating operating system values
- Checking compliance fields
- Removing invalid records where appropriate

The objective is to create a consistent dataset that can be used for reliable downstream analysis.

---

## 9. Data Transformation

SQL and data transformation techniques are used to convert raw infrastructure information into an analytical dataset.

Important SQL concepts used in the project include:

- `CASE WHEN`
- `UNION ALL`
- Filtering
- Aggregation
- Conditional classification

### CASE WHEN

`CASE WHEN` can be used to classify records according to compliance conditions.

Example:

```sql
CASE
    WHEN compliance_status = 'Non-Compliant'
    THEN 'High Risk'
    ELSE 'Low Risk'
END
```

The exact classification logic depends on the project dataset and the defined business rules.

### UNION ALL

`UNION ALL` can be used to combine records from multiple compatible sources.

Example:

```sql
SELECT *
FROM source_1

UNION ALL

SELECT *
FROM source_2;
```

This allows multiple datasets to be combined into a unified analytical dataset.

---

## 10. Risk Classification

The analytical prediction dataset contains a field called:

```text
predicted_high_risk
```

The field is represented as:

```text
0 = Not classified as high-risk
1 = Classified as high-risk
```

This field is used by both the Power BI and Streamlit dashboards.

The output allows the dashboard to calculate:

```text
Total Servers
High-Risk Servers
High Risk Rate
```

The classification output can then be analyzed across different infrastructure dimensions.

---

## 11. Risk Analysis Logic

The dashboard calculates the total number of analyzed servers using:

```text
Total Servers = Count of server_id
```

The number of high-risk servers is calculated using:

```text
High-Risk Servers =
Count of server_id
where predicted_high_risk = 1
```

The high-risk rate is calculated using:

```text
High Risk Rate =
High-Risk Servers / Total Servers × 100
```

For example, if:

```text
Total Servers = 1,000
High-Risk Servers = 227
```

then:

```text
High Risk Rate = 227 / 1,000 × 100
               = 22.70%
```

The displayed values change when filters are applied.

---

## 12. Power BI Dashboard

The Power BI dashboard is titled:

```text
CDC Infrastructure Compliance & High-Risk Server Dashboard
```

The dashboard provides an interactive view of infrastructure risk.

The Power BI data model contains a table:

```text
ml_predictions
```

Important fields include:

```text
server_id
predicted_high_risk
priority
operating_system
environment
incident_category
```

A measure called:

```text
High Risk Rate
```

is also used in the dashboard.

---

## 13. Power BI Dashboard KPIs

The dashboard contains three primary metrics.

### Test Set Servers

Represents the number of servers included in the analyzed dataset.

### Predicted High-Risk

Represents the number of servers where:

```text
predicted_high_risk = 1
```

### High Risk Rate

Represents the percentage of analyzed servers classified as high-risk.

Formula:

```text
High Risk Rate =
High-Risk Servers / Total Servers × 100
```

---

## 14. Power BI Dashboard Visualizations

The Power BI dashboard contains the following major visualizations.

### Predicted High-Risk Servers by Priority

Shows the distribution of high-risk servers across different priority levels.

This helps analysts understand the concentration of high-risk servers by priority.

### Predicted High-Risk Servers by Operating System

Shows the distribution of high-risk servers across operating systems.

For example:

```text
Windows
Red Hat Enterprise Linux
```

This allows analysts to identify operating-system-specific risk patterns.

### Predicted High-Risk Servers by Environment

Shows the number of high-risk servers across environments such as:

```text
Production
Development
Test
```

This helps analysts understand where high-risk infrastructure is concentrated.

### Predicted High-Risk by Incident Category

Shows high-risk servers grouped by incident category.

This helps identify categories associated with a larger number of high-risk records.

---

## 15. Streamlit Dashboard

A Streamlit dashboard was developed to make the project available as an interactive Python-based web application.

The Streamlit dashboard provides:

- Interactive filters
- KPI cards
- Interactive charts
- Filtered data tables
- High-risk calculations
- Server-level analysis

The application is developed using:

```text
Python
Pandas
Plotly
Streamlit
OpenPyXL
```

---

## 16. Streamlit Dashboard Title

The dashboard is titled:

```text
CDC Infrastructure Compliance & High-Risk Server Dashboard
```

---

## 17. Streamlit Dashboard Filters

The dashboard provides four main filters.

### Environment

Allows users to select specific environments.

### Operating System

Allows users to filter by operating system.

### Priority

Allows users to filter by priority.

### Incident Category

Allows users to filter by incident category.

The selected filters are applied simultaneously to the dataset.

---

## 18. Streamlit Dashboard KPIs

The Streamlit dashboard displays three primary KPI cards:

```text
Test Set Servers
Predicted High-Risk
High Risk Rate
```

The KPI values change dynamically based on the selected filters.

For example:

```text
Test Set Servers: 1,000
Predicted High-Risk: 227
High Risk Rate: 22.70%
```

The actual displayed values depend on the selected filters and dataset.

---

## 19. Streamlit Dashboard Visualizations

The Streamlit dashboard contains four primary charts.

### 1. Predicted High-Risk Servers by Priority

A bar chart showing high-risk server counts by priority.

### 2. Predicted High-Risk Servers by Operating System

A pie chart showing the distribution of high-risk servers across operating systems.

### 3. Predicted High-Risk Servers by Environment

A bar chart showing high-risk server counts by environment.

### 4. Predicted High-Risk by Incident Category

A horizontal bar chart showing high-risk server counts by incident category.

---

## 20. Filtered Data Summary

The dashboard also provides a filtered data table.

The table changes based on the selected filters.

The dashboard displays:

```text
Showing X records after applying the selected filters.
```

This allows users to inspect the records contributing to the dashboard metrics.

---

## 21. Python Application Workflow

The Streamlit application follows this workflow:

```text
Load Excel Dataset
       |
       v
Clean Column Names
       |
       v
Validate Required Columns
       |
       v
Convert Data Types
       |
       v
Apply Dashboard Filters
       |
       v
Calculate KPIs
       |
       v
Create Risk Dataset
       |
       v
Generate Charts
       |
       v
Display Filtered Data
```

---

## 22. Required Dataset Columns

The Streamlit application expects the following columns:

```text
server_id
predicted_high_risk
priority
operating_system
environment
incident_category
```

The application validates whether these columns exist before generating the dashboard.

If a required column is missing, the application displays an error and lists the available columns.

---

## 23. Repository Structure

The GitHub repository contains:

```text
it-compliance-risk-analytics-dashboard/
│
├── README.md
│
├── app.py
│
├── requirements.txt
│
├── ml_predictions.xlsx
│
└── tcs_cdc_raw_compliance.csv
```

---

## 24. File Descriptions

### README.md

Contains project documentation, architecture, workflow, dashboard explanation, interview preparation, and learning notes.

### app.py

The main Streamlit application.

It contains the dashboard logic including:

- Data loading
- Data validation
- Filtering
- KPI calculation
- Risk analysis
- Plotly visualizations
- Filtered data table

### requirements.txt

Contains the Python packages required to run the Streamlit application.

```text
streamlit
pandas
plotly
openpyxl
```

### ml_predictions.xlsx

Contains the prediction/risk output used by the Streamlit dashboard.

Important fields include:

```text
server_id
predicted_high_risk
priority
operating_system
environment
incident_category
```

### tcs_cdc_raw_compliance.csv

Contains the raw infrastructure compliance dataset used as the source data for the project.

---

## 25. Technology Stack

### Programming

- Python
- SQL

### Data Analysis

- Pandas
- NumPy
- Excel
- MySQL

### Visualization

- Power BI
- Plotly
- Matplotlib
- Seaborn

### Business Intelligence

- Power BI
- DAX

### Web Dashboard

- Streamlit

### Version Control

- GitHub

### Deployment

- Streamlit Community Cloud

---

## 26. Why Power BI and Streamlit Both?

Power BI and Streamlit serve different purposes in this project.

### Power BI

Power BI is useful for:

- Business intelligence
- KPI reporting
- Interactive dashboards
- DAX calculations
- Business-user analysis
- Enterprise reporting

### Streamlit

Streamlit is useful for:

- Python-based dashboards
- Portfolio demonstrations
- Interactive web applications
- Deploying data applications
- Demonstrating Python and data analytics skills

Using both tools demonstrates the ability to present analytical results through both BI and Python-based environments.

---

## 27. Deployment Architecture

The Streamlit deployment follows this architecture:

```text
GitHub Repository
       |
       v
Streamlit Community Cloud
       |
       v
Python Application
       |
       v
ml_predictions.xlsx
       |
       v
Interactive Web Dashboard
```

When the application is deployed, Streamlit Community Cloud runs the Python application and makes the dashboard accessible through a web URL.

---

## 28. Live Dashboard

The deployed Streamlit dashboard is available here:

[https://it-compliance-risk-analytics-dashboard-nl9odxydhhaz7avm5gbdyi.streamlit.app/](https://it-compliance-risk-analytics-dashboard-nl9odxydhhaz7avm5gbdyi.streamlit.app/)

---

## 29. GitHub Repository

GitHub repository:

[https://github.com/shubham140300/it-compliance-risk-analytics-dashboard](https://github.com/shubham140300/it-compliance-risk-analytics-dashboard)

---

## 30. How to Run the Project Locally

### Step 1 - Clone the Repository

```bash
git clone https://github.com/shubham140300/it-compliance-risk-analytics-dashboard.git
```

### Step 2 - Navigate to the Project

```bash
cd it-compliance-risk-analytics-dashboard
```

### Step 3 - Install Dependencies

```bash
pip install -r requirements.txt
```

### Step 4 - Run Streamlit

```bash
streamlit run app.py
```

### Step 5 - Open the Dashboard

Streamlit will provide a local URL, normally similar to:

```text
http://localhost:8501
```

Open the URL in a browser.

---

## 31. How to Use the Dashboard

### Step 1

Open the Streamlit dashboard.

### Step 2

Use the sidebar filters.

Available filters:

```text
Environment
Operating System
Priority
Incident Category
```

### Step 3

Select or remove filter values.

### Step 4

Observe the KPI cards.

The dashboard dynamically updates:

```text
Test Set Servers
Predicted High-Risk
High Risk Rate
```

### Step 5

Analyze the charts.

Check:

- Priority distribution
- Operating system distribution
- Environment distribution
- Incident category distribution

### Step 6

Review the filtered data table.

The table shows the records remaining after the selected filters are applied.

---

## 32. Example Analysis

Suppose the dashboard displays:

```text
Test Set Servers: 1,000
Predicted High-Risk: 227
High Risk Rate: 22.70%
```

This means:

```text
Total analyzed servers = 1,000

Servers classified as high-risk = 227

High-risk rate = 22.70%
```

An analyst can then use the charts to investigate:

```text
Which priority has the highest number of high-risk servers?
Which operating system has more high-risk servers?
Which environment has more high-risk servers?
Which incident category has more high-risk records?
```

The filters can then be used to investigate specific segments.

---

## 33. Business Use Case

This type of dashboard can support infrastructure and compliance teams in:

- Monitoring compliance status
- Identifying potentially high-risk infrastructure
- Prioritizing investigation
- Identifying recurring incident categories
- Understanding risk distribution
- Monitoring infrastructure segments
- Supporting remediation activities
- Creating management-level reports

The dashboard provides an analytical view of the available data and does not replace the organization's security or compliance controls.

---

## 34. Example Business Questions

The dashboard can help answer questions such as:

### Question 1

How many servers are currently classified as high-risk?

Use the `Predicted High-Risk` KPI.

### Question 2

What percentage of analyzed servers are high-risk?

Use the `High Risk Rate` KPI.

### Question 3

Which priority contains the highest number of high-risk servers?

Use the `Predicted High-Risk Servers by Priority` chart.

### Question 4

Which operating system has more high-risk servers?

Use the `Predicted High-Risk Servers by Operating System` chart.

### Question 5

Which environment contains more high-risk servers?

Use the `Predicted High-Risk Servers by Environment` chart.

### Question 6

Which incident category is associated with more high-risk servers?

Use the `Predicted High-Risk by Incident Category` chart.

---

## 35. Interview Explanation

A concise way to explain the project during an interview is:

> I built an IT Compliance and Risk Analytics Dashboard to analyze infrastructure compliance data and identify potentially high-risk servers. I worked with server-level compliance attributes, transformed the data using SQL and Python-based processing, and used a predicted high-risk flag for risk segmentation. I then created a Power BI dashboard with KPIs and risk analysis by priority, operating system, environment, and incident category. To make the project accessible as a portfolio application, I also built a Streamlit dashboard using Python, Pandas, and Plotly and deployed it through Streamlit Community Cloud.

---

## 36. Technical Interview Questions

### Q1. What was the purpose of this project?

The purpose was to analyze infrastructure compliance data, identify potentially high-risk servers, and create interactive dashboards for risk monitoring and analysis.

### Q2. What data did you use?

The project uses infrastructure compliance data containing server-level information and compliance/security indicators.

### Q3. Why did you use SQL?

SQL was used for data retrieval, transformation, filtering, combining records, aggregation, and classification logic.

### Q4. Why did you use Power BI?

Power BI was used to create interactive business intelligence dashboards, KPIs, DAX calculations, and visual analysis.

### Q5. Why did you use Streamlit?

Streamlit allowed the project to be converted into an interactive Python-based web dashboard that can be accessed through a browser.

### Q6. What is `predicted_high_risk`?

`predicted_high_risk` is the output field used to identify whether a server record is classified as high-risk.

```text
0 = Not high-risk
1 = High-risk
```

### Q7. How did you calculate High Risk Rate?

The calculation is:

```text
High Risk Rate =
High-Risk Servers / Total Servers × 100
```

### Q8. How does filtering work in Streamlit?

The application uses Streamlit multiselect filters.

The selected values are used to filter the Pandas DataFrame.

For example:

```python
filtered_df = df[
    df["environment"].isin(environment_filter)
    &
    df["operating_system"].isin(os_filter)
    &
    df["priority"].isin(priority_filter)
    &
    df["incident_category"].isin(incident_filter)
].copy()
```

### Q9. Why did you use Pandas?

Pandas was used for:

- Data loading
- Data cleaning
- Data transformation
- Filtering
- Grouping
- Aggregation
- KPI calculations

### Q10. Why did you use Plotly?

Plotly was used to create interactive visualizations for the Streamlit dashboard.

### Q11. How did you deploy the dashboard?

The Streamlit application was stored in GitHub and deployed using Streamlit Community Cloud.

The deployment architecture is:

```text
GitHub
   |
   v
Streamlit Community Cloud
   |
   v
app.py
   |
   v
Interactive Dashboard
```

---

## 37. Important Technical Concepts to Study

To confidently explain this project in an interview, study the following concepts.

### SQL

Study:

- SELECT
- WHERE
- GROUP BY
- ORDER BY
- CASE WHEN
- UNION ALL
- JOIN
- Aggregate functions
- COUNT
- SUM
- NULL handling
- Subqueries
- CTEs

### Python

Study:

- Variables
- Lists
- Dictionaries
- Functions
- Loops
- Conditional statements
- Exception handling
- File handling

### Pandas

Study:

- DataFrame
- Series
- `read_csv()`
- `read_excel()`
- `head()`
- `info()`
- `describe()`
- `isnull()`
- `fillna()`
- `dropna()`
- `drop_duplicates()`
- `groupby()`
- `merge()`
- `concat()`
- Filtering
- `sort_values()`

### Power BI

Study:

- Power Query
- Data modeling
- Relationships
- Star schema
- Calculated columns
- Measures
- DAX
- Filters
- Slicers
- KPI cards
- Bar charts
- Pie charts
- Drill-down
- Dashboard design

### DAX

Important concepts include:

- `CALCULATE`
- `COUNT`
- `COUNTROWS`
- `SUM`
- `DIVIDE`
- `FILTER`
- `DISTINCTCOUNT`
- Variables
- Filter context
- Row context

### Streamlit

Study:

- `st.set_page_config()`
- `st.title()`
- `st.header()`
- `st.subheader()`
- `st.write()`
- `st.metric()`
- `st.sidebar`
- `st.multiselect()`
- `st.dataframe()`
- `st.cache_data`
- Streamlit application flow

### Plotly

Study:

- Bar charts
- Pie charts
- Horizontal bar charts
- Figure configuration
- Interactive filtering
- Layout customization

---

## 38. Project Limitations

This project has several limitations.

### 1. Data Dependency

Dashboard results depend on the quality and completeness of the underlying dataset.

### 2. Prediction Dependency

The risk analysis depends on the `predicted_high_risk` output.

The dashboard itself does not independently validate the prediction model.

### 3. Static Prediction File

The current Streamlit dashboard reads the prediction dataset from:

```text
ml_predictions.xlsx
```

A production implementation could connect directly to a database or data pipeline.

### 4. No Real-Time Data Pipeline

The current portfolio implementation is not a real-time infrastructure monitoring system.

### 5. Portfolio Deployment

The public deployment is intended for portfolio demonstration and should use synthetic, anonymized, or otherwise approved data.

---

## 39. Future Improvements

Possible future improvements include:

### Real-Time Data Pipeline

Connect the dashboard to a database or streaming data source.

### Automated ETL

Build an automated pipeline for:

```text
Data Collection
      |
      v
Data Validation
      |
      v
Transformation
      |
      v
Prediction
      |
      v
Dashboard
```

### Machine Learning Model

Develop and evaluate a formal machine learning model for high-risk classification.

Potential models could include:

- Logistic Regression
- Decision Tree
- Random Forest
- Gradient Boosting
- XGBoost

The model should be evaluated using appropriate metrics such as:

- Accuracy
- Precision
- Recall
- F1-score
- ROC-AUC

### Model Monitoring

Future versions could monitor:

- Model performance
- Prediction drift
- Data drift
- Feature drift
- False positives
- False negatives

### Database Integration

Instead of loading Excel files, the application could connect to:

```text
MySQL
PostgreSQL
AWS RDS
Cloud Data Warehouse
```

### Automated Refresh

A production version could automatically refresh compliance information and prediction results.

### Authentication

A production enterprise application could include authentication and role-based access.

### Alerting

Future versions could generate alerts when:

```text
High Risk Rate > Threshold
```

or when critical infrastructure becomes non-compliant.

---

## 40. Learning Notes

This project can be used as a study project for understanding an end-to-end analytics workflow.

The key learning sequence is:

```text
SQL
   ↓
Data Cleaning
   ↓
Data Transformation
   ↓
Risk Classification
   ↓
Data Analysis
   ↓
Power BI
   ↓
DAX
   ↓
Python
   ↓
Pandas
   ↓
Streamlit
   ↓
Plotly
   ↓
GitHub
   ↓
Cloud Deployment
```

---

## 41. What I Learned From This Project

Through this project, the following concepts can be practiced:

- Working with infrastructure datasets
- Data cleaning
- Data transformation
- SQL-based analysis
- Risk classification
- KPI calculation
- Data visualization
- Power BI dashboard development
- DAX calculations
- Python data analysis
- Pandas
- Plotly
- Streamlit
- GitHub
- Cloud deployment
- Dashboard design
- Business problem solving

---

## 42. Key Takeaways

The main lessons from the project are:

1. Raw infrastructure data needs to be cleaned before analysis.
2. SQL is useful for data retrieval and transformation.
3. Risk classification makes large datasets easier to analyze.
4. KPIs provide a quick overview of the dataset.
5. Visualization helps identify patterns and concentrations of risk.
6. Power BI is useful for business intelligence reporting.
7. Python and Streamlit can be used to build interactive analytical applications.
8. GitHub provides version control and project visibility.
9. Streamlit Community Cloud can be used to deploy Python dashboards.
10. A complete analytics project should connect technical implementation with a clear business problem.

---

## 43. Project Links

### GitHub Repository

[https://github.com/shubham140300/it-compliance-risk-analytics-dashboard](https://github.com/shubham140300/it-compliance-risk-analytics-dashboard)

### Live Streamlit Dashboard

[https://it-compliance-risk-analytics-dashboard-nl9odxydhhaz7avm5gbdyi.streamlit.app/](https://it-compliance-risk-analytics-dashboard-nl9odxydhhaz7avm5gbdyi.streamlit.app/)

---

## 44. Data Disclaimer

This project is intended for portfolio, educational, and demonstration purposes.

Any public version of the project should use synthetic, anonymized, or otherwise approved data.

Confidential enterprise information, real production server identifiers, internal IP addresses, credentials, security information, or proprietary company data should not be exposed through a public GitHub repository or public dashboard.

The public dashboard should therefore contain only data that is safe and authorized for external demonstration.

---

## 45. Final Project Summary

This project demonstrates an end-to-end IT compliance and risk analytics workflow:

```text
Infrastructure Compliance Data
            |
            v
       Data Cleaning
            |
            v
      SQL Transformation
            |
            v
     Risk Classification
            |
            v
      Analytical Dataset
            |
      +-----+------+
      |            |
      v            v
   Power BI     Streamlit
   Dashboard    Dashboard
      |            |
      v            v
 Business       Interactive
 Insights       Web Analytics
```

The project combines data analytics, SQL, Python, Power BI, DAX, visualization, GitHub, and cloud deployment into one portfolio solution.

It can be used to demonstrate practical knowledge of data analysis, business intelligence, dashboard development, and an end-to-end analytics workflow.
