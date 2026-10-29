# bank-reconciliation-tool
\
A Python tool that compares a company's ledger against its bank statement\
and automatically finds every difference.\
\
## The problem\
Finance teams check their own records against the bank's records to make\
sure they match. This is called reconciliation, and it's often done by hand\
in Excel, which is slow and easy to get wrong.\
\
## What it does\
The tool finds four types of problem:\
- Transactions in the ledger but not on the bank statement\
- Transactions on the bank statement but not in the ledger\
- Duplicate charges\
- Transactions where the amounts don't match\
\
It then proves the reconciliation is correct by applying every difference\
to the ledger total and checking it lands exactly on the bank total.\
Results are saved to an Excel report with one sheet per problem type.\
\
## Results\
Tested on practice data containing 8 planted errors. The tool found all 8\
and reconciled to a difference of \'a30.00.\
\
## How to run it\
1. Install the libraries: `pip install pandas openpyxl`\
2. Put `company_ledger.csv` and `bank_statement.csv` in the same folder\
3. Run `python3 reconcile.py`\
4. Open `reconciliation_report.xlsx`\
\
## Built with\
Python, pandas, openpyxl\
\
## Limitations and next steps\
- Currently matches transactions using a reference number. Real bank\
  statements often have none, so a future version will match on date,\
  amount and description.\
- Only handles two files at a time.\
\
## Note\
All data in this project is fake and generated for practice.}
