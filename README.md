Student Finance Tracker – Google Sheets and Apps Script

A simple template for students to record and manage their monthly finances using Google Sheets and Google Apps Script. Transactions are organized into separate monthly sheets, covering November 2026 through March 2027. The remaining balance at the end of each month is automatically carried over to the next month.

Features

- Monthly sheets: Separate sheets for `Nov 2026`, `Dec 2026`, `Jan 2027`, `Feb 2027`, and `Mar 2027`.
- Automatic balance carryover: The previous month's closing balance is added as the first income entry, labeled "Previous Month's Balance".
- Transaction categories: Category options change based on the selected transaction type: income or expense.
- Custom categories: Add or edit income and expense categories through the `Pengaturan` sheet.
- Monthly summaries: View the opening balance, total income, total expenses, closing balance, financial status, expense ratio, and transaction count.
- Running balance: Track the balance after each transaction.
- Category breakdown: View transaction totals by category on each monthly sheet.
- Overall summary: The `Rekap` sheet combines data from all months and includes an income-versus-expenses chart.
- Input validation: Checks whether transaction dates match the selected month and prevents negative transaction amounts.
- Automatic formatting: Income entries are highlighted in green, expenses in red, and negative balances are marked in red.

The first row of the transaction table is filled in automatically and should not be edited manually.

Installation

1. Create a new spreadsheet in Google Sheets.
2. Open **Extensions → Apps Script**.
3. Clear the existing code in `Code.gs`, then copy the entire contents of [`keuangan_mahasiswa.gs`](https://claude.ai/chat/keuangan_mahasiswa.gs).
4. Paste the code into the Apps Script editor and save it.
5. Select the `setupKeuangan` function and click **Run**.
6. Follow the prompts to grant the required permissions.
7. Return to your spreadsheet. The required sheets will be created automatically, and the **Keuangan** menu will become available.

If the menu does not appear, refresh the spreadsheet page.

How to Use

1. Open the `Pengaturan` sheet. If you have money left over before November 2026, enter your starting balance in cell `D2`. You can also adjust the transaction categories as needed.
2. Open the monthly sheet you want to use, such as `Nov 2026`.
3. Start entering transactions from row 10 and fill in the following columns:
   - Date: The transaction date, which must match the month of the sheet.
   - Type: Select `Pemasukan` (Income) or `Pengeluaran` (Expense).
   - Category: Choose an option from the dropdown list.
   - Description: Add a note about the transaction.
   - Amount: Enter a positive amount without a minus sign.
4. Check the financial summary at the top of the sheet or open `Rekap` to review the totals for all months.
To recreate a monthly sheet, delete that sheet first and run `setupKeuangan` again.

Technologies
- Google Sheets
- Google Apps Script (JavaScript)

Contributing
Suggestions and improvements are welcome. If you find a bug or have an idea for a new feature, feel free to open an issue or submit a pull request.

