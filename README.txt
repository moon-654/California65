Rule-Based Prop 65 IMDS Report Generator

How to run
1. Install Python 3.10+
2. pip install -r requirements.txt
3. streamlit run app.py

What is included
- app.py
- requirements.txt
- prop65_rules_template.xlsx

Why this version is better
- Rules are externalized in Excel
- Chemicals can be added without editing code
- Special logic such as Lithium-carbonate follow-up can be maintained in the rules sheet

Important limitation
- Output quality depends on the PDF having extractable text and a reasonably similar IMDS table structure
- Final external use still requires technical review
