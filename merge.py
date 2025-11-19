import zipfile
import pandas as pd
import os

# ----  Extract ZIP file ----
zip_path = r"C:\Users\kemer\Downloads\All Datasets.zip"
extract_folder = "extracted_files"

with zipfile.ZipFile(zip_path, 'r') as zip_ref:
    zip_ref.extractall(extract_folder)

# ---- Find all Excel files recursively ----
excel_files = []
for root, dirs, files in os.walk(extract_folder):
    for file in files:
        if file.endswith((".xlsx", ".xls")):
            excel_files.append(os.path.join(root, file))

if len(excel_files) == 0:
    raise Exception("No Excel files found in the ZIP folder or its subdirectories.")

print(f"Found {len(excel_files)} Excel files:")
for f in excel_files:
    print(f"  - {f}")

# ---- Load , clean and merge ----
dfs = []

rename_map = {
    # consultant related
    "CONSULTANT": "Consultant",
    "consultant": "Consultant",
    "Job founder": "Consultant",
    "job founder": "Consultant",
    "Job founder": "Consultant",
    "Consultent": "Consultant",

    # source/url related
    "URL": "URLSource",
    "Url": "URLSource",
    "url": "URLSource",
    "Source": "URLSource",
    "Sourse": "URLSource",
    "source": "URLSource",
    "Source of the advertisement": "URLSource",
    "source of the advertisement": "URLSource",
    "Source (weblink,newspaper..)": "URLSource",
    "URL": "URLSource",

    # job title related
    "Job_title": "Job_Title",
    "JobTitle": "Job_Title",
    "JOB TITLE": "Job_Title",
    "Job Title": "Job_Title",
    "Job_Title": "Job_Title",
    "Job role": "Job_Title",
    "Role(job name)": "Job_Title",
    "Job Category": "Job_Title",
    "Job category": "Job_Title",

    # Location related
    "LOCATION": "Location",
    "Location": "Location",


    # Country related
    "location(country)": "Country",
    "Country": "Country",

    # ID related
    "Id": "ID",
    "ID": "ID",

    # Date Retrieved related
    "DateRetrieved": "DateRetrieved",
    "DATE RETRIEVED": "DateRetrieved",
    "Date_Retrieved": "DateRetrieved",

    # Company related
    "COMPANY": "Company",
    "Company or Organization name": "Company",
    "Company": "Company",

    # Experience Category related
    "Experience_Category": "Experience_Category",
    "Job Experience":"Experience_Category",
    "Experiance category":"Experience_Category",
    "Experience": "Experience_Category",
    "year's experience": "Experience_Category",
    "Experience": "Experience_Category",
    "Experience": "Experience_Category",
    "Job Experience": "Experience_Category",
    "Experience_Category": "Experience_Category",
    "MinimumExperience": "Experience_Category",
    "Experience_Area":"Experience_Category",

    # Mode related
    "Mode": "Mode",
    "Model":  "Mode",
    "Model(online,offline)": "Mode",
    "Employment_Type": "Mode",
    "Mode": "Mode",
    "Work TYPE (ON SITE/ REMOTE / HYBRID)": "Mode",
    "WorkMode": "Mode",
    "Hybrid": "Mode",

    # Knowledge in related
    "Knowledge_in": "Knowledge_in",
    "Knowledge in": "Knowledge_in", 

    # Educational Qualifications related
    "Educational_qualifications": "Educational_Qualifications", 
    "Educational_qualifications": "Educational_Qualifications", 
    "Educational Qualification": "Educational_Qualifications", 
    "Education requirement": "Educational_Qualifications", 
    "Education qualifications": "Educational_Qualifications", 
    "Educational qualifications": "Educational_Qualifications", 
    "Education requirement": "Educational_Qualifications", 
    "Educational_qualifications": "Educational_Qualifications", 

    # Salary related
    "Salary": "Salary",
    " Salary": "Salary",
    "Salary_per_month(Rs)": "Salary",

    # English needed related
    "English_needed": "English Needed",
    "English Needed": "English Needed",
    "English needed": "English Needed",
    "English knowledge": "English Needed",
    "Languages": "English Needed",
    "English knowledge skills": "English Needed",
    "ENGLISH": "English Needed",

    # BSc needed related
    "BSc_needed": "BSC_Needed",
    "BSc_needed": "BSC_Needed",
    "BSc needed": "BSC_Needed",
    "BSC_Needed": "BSC_Needed",
    "Bsc.": "BSC_Needed",
    "BScNeeded": "BSC_Needed",
    "BSc_needed": "BSC_Needed",

    # MSc needed related
    "MSc_needed": "MSc_needed",
    "MSc_needed": "MSc_needed",
    "MSc needed": "MSc_needed",
    "MSC_Needed": "MSc_needed",
    "MSC": "MSc_needed",
    "MScNeeded": "MSc_needed",
    "MSc_needed": "MSc_needed",

    # PhD needed related
    "PhD_needed": "PhD_needed",
    "PhD_Needed": "PhD_needed",
    "phD_needed": "PhD_needed",
    "PhD needed": "PhD_needed",
    "PHD": "PhD_needed",
    "PhDNeeded": "PhD_needed",

    # Date Published related
    "DatePublished": "Date_Published",
    "Date posted": "Date_Published",
    "Date_Published": "Date_Published",
    "Date Published": "Date_Published",
    "DATE PUBLISHED": "Date_Published",
    "DatePublished": "Date_Published",
    "DatePublished": "Date_Published",

    # Salary related
    "Salary": "Salary",
    "Salary_per_month(Rs)": "Salary",

    # Slill related
    "Ms Word": "MS_Word",
    "MS_Word": "MS_Word",
    "MS Word": "MS_Word",
    "MS word": "MS_Word",
    "MS WORD": "MS_Word",

    "Ms PowerPoint": "MS_PowerPoint",
    "MS_PowerPoint": "MS_PowerPoint",
    "Ms Powerpoint": "MS_PowerPoint",
    "MS power point": "MS_PowerPoint",
    "MS POWER POINT": "MS_PowerPoint",

    "TEAMS": "Teams",
    "Teams": "Teams",

    "PYTHON": "Python",
    "Python": "Python",

    "JAVA": "Java",
    "Java": "Java",

    "SCALA": "Scala",
    "Scala": "Scala",

    "OOP( OBJECTED- ORIENTED PROGRAMMING)": "OOP Concepts with applications",
    "OOP Concepts with applications": "OOP Concepts with applications",

    "Data_warehouse": "Data_warehouse",
    "Data_Warehousing":"Data_warehouse",
    "Data warehousing":"Data_warehouse",
    "Data warehouses":"Data_warehouse",
    "Warehouse Architectures": "Data_warehouse",

    "Familiar with GitHub": "Git",
    "Git": "Git",
    "Github": "Git",
    "GIT": "Git",

    "Data analyst role": "Analytical_Skill",
    "Analytical_Skill": "Analytical_Skill",
    "AnalysisSkills": "Analytical_Skill",
    "Analytical Skills": "Analytical_Skill",
    "Analytical & Thinking Skills": "Analytical_Skill",


    "Organizational skills": "Organizing_Skills",
    "Organizing_Skills": "Organizing_Skills",
    "Organization & Time Management": "Organizing_Skills",

    "Google_Cloud": "Cloud",
    "Cloud": "Cloud",
    "Google Cloud": "Cloud",
    "Cloud Data Service": "Cloud",
    "Google Cloud(GCP)": "Cloud",

    "* Exploratory Data Analysis (EDA)": "EDA_Experience",
    "EDA_Experience": "EDA_Experience",

    " PySpark": "Pyspark",
    "Pyspark": "Pyspark",
    "pyspark": "Pyspark",

    "Databricks": "DataBricks",
    " Databricks": "DataBricks",
    "DataBricks": "DataBricks",

    "BigQuery Stack": "BigQuery",
    "BigQuery": "BigQuery",

    "No SQL": "NoSQL",
    "NoSQL": "NoSQL",

    "MS SQL": "MS_SQL",
    "MS_SQL": "MS_SQL",

    "Data Cleaning": "Data Cleaning",
    "Data_cleaning": "Data Cleaning",
    "Data_cleaning": "Data Cleaning",

    "collaboration": "Collaboration",
    "Collaboration": "Collaboration",
    "Collaborating": "Collaboration",
    "Teamwork & Collaboration": "Collaboration",

    "Reporting Skills": "Data handling and reporting",
    "Data handling and reporting": "Data handling and reporting",
    "* Report Preparation": "Data handling and reporting",

    "Canva Design": "CANVA",
    "CANVA": "CANVA",

    "Bayesian methodology": "Bayesian",
    "Bayesian": "Bayesian",

    "Policy-level data interpretation": "Interpretation skills",
    "Interpretation skills": "Interpretation skills",

    "looker Studio": "Looker",
    "Looker": "Looker",
    " Looker (data visualization tools)": "Looker",

    "Data structures": "Data Structures",
    "Data Structures": "Data Structures",

    "Web Aanalytics tools": "Web Aanalytics tools",
    "Web_Analytic_tools": "Web Aanalytics tools",

    "Pandas": "Pandas",
    "Pandas": " Pandas",

    "Problem solving": "Problem_Solving_Skills",
    "Problem_Solving_Skills": "Problem_Solving_Skills",
    "Problem_Solving": "Problem_Solving_Skills",
    "ProblemSolving and Analytical Skills": "Problem_Solving_Skills",
    "Analytical /Problem solving skills": "Problem_Solving_Skills",

    "Data governance": "Data_governance",
    "Data_governance": "Data_governance",
    "Data Governance & Security": "Data_governance",

    "Team_Handling": "Teamwork",
    "TeamWorkAbiilities": "Teamwork",
    "Teamwork": "Teamwork",
    "Team work": "Teamwork",


    "Ms Excel": "MS_Excel",
    "MS_Excel": "MS_Excel",
    "Excel": "MS_Excel",
    "MS Excel": "MS_Excel",
    "MS EXCEL ": "MS_Excel",

    "MS ACCESS": "MS_Access",
    "MS_Access": "MS_Access",

    "Data Visualization": "Data_Visualization",
    "Data_Visualization": "Data_Visualization",
    "Data_visualization": "Data_Visualization",
    "DataVisualization": "Data_Visualization",
    "Data visualization": "Data_Visualization",

    "Machine Learning": "ML",
    "ML": "ML",
    "Machine_Learning": "ML",
    "ML Models": "ML",
    "Machine Learnning": "ML",
    "Machine Learning (principles/systems.technologies,techniques)": "ML",

    "finance-sector concepts": "Finance_Knowledge",
    "Finance_Knowledge": "Finance_Knowledge",

    "Communication Skills": "Communication_Skills",
    "Communication_Skills": "Communication_Skills",
    "CommunicationSkills": "Communication_Skills",
    "Communication & Interpersonal Skills": "Communication_Skills",
    "Communication": "Communication_Skills",
    "Communication skills": "Communication_Skills",

    "NumPy": "NumPy",
    "Numpy": "NumPy",

    "Power_BI": "PowerBI",
    "PowerBI": "PowerBI",
    "Microsoft Power BI": "PowerBI",
    "Power BI": "PowerBI",
    " Power BI (BI tools)": "PowerBI",

    "Tableau (BI tools)": "Tableau",
    "Tableau": "Tableau",

    "Data_mining": "Data_mining",
    "Data Mining": "Data_mining",

    "BigData": "BigData",
    "Big_Data": "BigData",
    "Big Data": "BigData",

    "Leadership": "Leadership_Skills",
    "Leadership_Skills": "Leadership_Skills",

    "Data Pipeline(ETL/ELT)": "Data_Pipelines",
    "Data_Pipelines": "Data_Pipelines",
    "Data Pipelines & Quality Monitoring": "Data_Pipelines",
    "Data Pipeline": "Data_Pipelines",
    "* ML pipelines (end-to-end)": "Data_Pipelines",

    "Natural Language Processing(NLP)": "Natural_Language_Processing(NLP)",
    "Natural_Language_Processing(NLP)": "Natural_Language_Processing(NLP)",
    "* NLP fundamentals": "Natural_Language_Processing(NLP)",

    "Presentattion Skills": "Presentation_Skills",
    "Presentation_Skills": "Presentation_Skills",
    "Presentation Skills": "Presentation_Skills",
    "Presentation skills": "Presentation_Skills",

    "Linux": "Unix/Linux",
    "Unix/Linux": "Unix/Linux",
    "Linux/Unix": "Unix/Linux",
    " Linux": "Unix/Linux",

    "statistical tests": "Statistical_Knowledge",
    "Statistical_Knowledge": "Statistical_Knowledge",
    "Statistical analysis": "Statistical_Knowledge",

    "statistical modelling": "statistical modelling",
    "Statistical Models": "statistical modelling",
    "Data Storage & Modeling": "statistical modelling",
    "Predictive & Modeling Techniques": "statistical modelling",

    "Advanced Multivariate & Exploratory Analysis": "Multivariate analysis",
    "Multivariate analysis": "Multivariate analysis",

    "Experimentation & Optimization": "Optimization",
    "Optimization": "Optimization",
    "* Model optimization (runtime, memory, accuracy)": "Optimization",

    "kafka": "Kafka",
    "Kafka": "Kafka",

    "Stata": "STATA",
    "STATA": "STATA",

    "Epidemiology": "Epidemiology",
    "Background in epidemiology": "Epidemiology",

    "JavaScript": "JavaScript",
    "Javascript": "JavaScript",

    "ETL tools": "ETL tools",
    "ETL": "ETL tools",
    "ETL Process Knowledge": "ETL tools",
    "ETL tools": "ETL tools",
    "ELT / ETL Framework": "ETL tools",

    "Google Analytics": "Google Analytics",
    "Google_Analytics": "Google Analytics",

    "Deep Learning(TensorFlow/PyTorch)": "Deep Learning",
    "Deep Learning": "Deep Learning",

    "Experience_Data Modelling": "Data_modeling",
    "Data_modeling": "Data_modeling",
    "Data modelling": "Data_modeling",
    "Data Modeling": "Data_modeling",

    "Experience working with large,complex datasets": "Large-scale data",
    "Large-scale data": "Large-scale data",

    "LLM": "LLM",
    "LLMs": "LLM",
    "* Large Language Models (LLMs)": "LLM",

    "Data_Management": "Data_Management",
    "Data_management": "Data_Management",

    "Scikit learn": "Scikit",
    "Scikit": "Scikit",
    " Scikit-learn (SKlearn)": "Scikit",

    "Time Series Analysis": "Time Series Analysis",
    "Time series forecasting": "Time Series Analysis",

     # Source related
    "Source_URL":"Source_Of_URL",
    "URL_Source":"Source_Of_URL",
    "Scource_of-URL":"Source_Of_URL",

    # Payment Frequency related
    "Payment Frequency": "Payment Frequency",
    "Payement_Freq": "Payment Frequency",




    
}

for file in excel_files:
    try:
        df = pd.read_excel(file)

        # ---- General column renaming rules ----
        df.rename(columns=lambda x: rename_map.get(x, x), inplace=True)

        # ---- SPECIAL RULE: ONLY for Group_6.xlsx ----
        if "Group_6.xlsx" in file:
            if "Source" in df.columns:
                df.rename(columns={"Source": "URLSource"}, inplace=True)

        # ---- Remove duplicate columns after renaming ----
        df = df.loc[:, ~df.columns.duplicated()]

        print(f"Loaded: {file} ({len(df)} rows)")
        dfs.append(df)

    except Exception as e:
        print(f"Error reading {file}: {e}")

if len(dfs) == 0:
    raise Exception("No Excel files could be loaded successfully.")

# ----  Merge all data ----
merged_df = pd.concat(dfs, ignore_index=True)

# ----  Save result ----
output_file = "merged_dataset.xlsx"
merged_df.to_excel(output_file, index=False)

print(f"\nMerge completed! File saved as {output_file}")
print(f"Total rows: {len(merged_df)}")
print(f"Total columns: {len(merged_df.columns)}")
