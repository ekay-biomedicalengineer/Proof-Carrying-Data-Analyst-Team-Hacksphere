# Project Title: Proof-Carrying Data Analyst (Agentic GenAI)
**Problem Statement Code:** HNX26PSI08  
**Team Name:** TEAM HACKSPHERE
**Repository Type:** Public  

---

## 1. What the Project Does
Our project is an AI-powered data analyst agent designed to answer complex questions across messy, multi-table datasets. 

Unlike traditional AI assistants that guess or hallucinate numerical answers, our system generates **verifiable, re-runnable Python code** for every calculation. It automatically handles real-world data issues (e.g., duplicate rows, missing values, mismatched currencies, and date ambiguities) and is programmed to explicitly refuse unanswerable or ambiguous questions whilst providing reasons as to why.

---

## 2. Core Logic & Reasoning Mechanism
The central architecture consists of three main components:
1. **Data Normalization & Inspection:** Scans multi-table datasets to identify schema mismatches, duplicate entries, missing values, and unit inconsistencies.
2. **Code Generation & Execution:** Translates natural language questions into executable Pandas/Python code to perform precise data analysis.
3. **Proof & Refusal Guardrails:** Runs the generated code in an execution pipeline to confirm accuracy. If the question is unanswerable due to missing data or inherent ambiguity, the agent refuses to guess and provides a reasoned explanation.

---

## 3. Data Pipeline
The flow of data through our system works as follows:
- **Collection:** Input data is provided via multiple CSV files (e.g., `sales.csv`, `customers.csv`, `refunds.csv`).
- **Processing:** Data tables are ingested, cleaned, and joined inside our Python execution environment.
- **Output Routing:** The processed data is queried by generated code, producing both the final numerical/text answer and the raw executable script as proof.

---

## 4. Scope Note (Minimally Viable Solution vs. Stretch Goals)

### Minimum Viable Solution (Implemented)
- Querying multi-table CSV datasets using LLM-generated Python code.
- Self-checking code execution to verify exact outputs.
- Handling basic messy data features (missing values, duplicate rows).
- Standard refusal reasoning for ambiguous/unanswerable questions.

### Stretch Goals / Future Work
- Dynamic auto-conversion for complex external APIs (e.g., live currency exchange rates).
- Interactive visual chart generation alongside code outputs.

---

## 5. Technologies, Libraries, and Models Used
- **Programming Language:** Python 3.10+
- **LLM / AI Framework:** [e.g., OpenAI API / LangChain / LlamaIndex]
- **Data Processing Libraries:** Pandas, NumPy
- **Environment:** Public Git Repository

---

## 6. How to Install Dependencies
To install and set up the required dependencies, run the following commands in your terminal:

```bash
# Clone the repository
git clone https://github.com/ekay-biomedicalengineer/Proof-Carrying-Data-Analyst-Team-Haccksphere.git

# Navigate into the project folder
cd Proof-Carrying-Data-Analyst-Team-Haccksphere

# Install required packages
pip install -r requirements.txt
```

---

## 7. How to Configure and Run the System
1. **Set Up API Keys:**
   Create a `.env` file in the root directory and add your API key:
   ```env
   OPENAI_API_KEY=your_api_key_here
   python main.py
