import pandas as pd
from pathlib import Path

# Find the folder this script is saved in
folder = Path(__file__).parent

# Load both files from that folder
ledger = pd.read_csv(folder / "company_ledger.csv")
bank = pd.read_csv(folder / "bank_statement.csv")

print("LEDGER (company records):")
print(ledger.head())
print()
print("BANK STATEMENT:")
print(bank.head())
print()

print("Number of ledger rows:", len(ledger))
print("Number of bank rows:", len(bank))
print()
print("Ledger total:", round(ledger["amount"].sum(), 2))
print("Bank total:", round(bank["amount"].sum(), 2))

missing_from_ledger = bank[~bank["ref"].isin(ledger["ref"])]
print("In the bank but NOT in the ledger:")
print(missing_from_ledger)

missing_from_bank = ledger[~ledger["ref"].isin(bank["ref"])]
print()
print("In the ledger but NOT in the bank:")
print(missing_from_bank)

duplicates = bank[bank.duplicated(subset="ref", keep=False)]
print()
print("Duplicates in the bank statement:")
print(duplicates)

# Put the ledger and bank side by side, matching on ref
both = pd.merge(ledger, bank, on="ref", suffixes=("_ledger", "_bank"))

# Keep only rows where the two amounts are different
different = both[both["amount_ledger"] != both["amount_bank"]]

print()
print("Amounts that don't match:")
print(different[["ref", "description_ledger", "amount_ledger", "amount_bank"]])

# The final proof: apply every fix to the ledger total
adjusted = ledger["amount"].sum()
adjusted = adjusted - missing_from_bank["amount"].sum()
adjusted = adjusted + missing_from_ledger["amount"].sum()
adjusted = adjusted + bank[bank.duplicated(subset="ref")]["amount"].sum()
adjusted = adjusted + (different["amount_bank"] - different["amount_ledger"]).sum()

print()
print("Ledger total after fixes:", round(adjusted, 2))
print("Bank total:", round(bank["amount"].sum(), 2))

# Save everything to one Excel file
output = folder / "reconciliation_report.xlsx"

with pd.ExcelWriter(output) as writer:
    missing_from_bank.to_excel(writer, sheet_name="Ledger only", index=False)
    missing_from_ledger.to_excel(writer, sheet_name="Bank only", index=False)
    duplicates.to_excel(writer, sheet_name="Duplicates", index=False)
    different.to_excel(writer, sheet_name="Amount mismatch", index=False)

print()
print("Report saved to:", output)