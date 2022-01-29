# Position-data-optimization


Open the *deliverable.py* when running. *function.py* is a function file.

Function input：
- **1Q21 position data:** drag the file directly into the terminal/enter the file path name;
- **Reporting period:** manual entry, no format requirements;
- **Select the position viewing method:** "Market value of holdings" or "Number of shares";
- **Market(multiple choices):** SSE/SESZ/SEHK;
- **Investment Type(multiple choices):** 1: REITs, 2: Passive Funds, 3: Stocks Long/Short, 4: Debt Hybrid Fund(Secondary Market), 5: Debt Hybrid Fund(Primary Market), 6: Dynamic Asset Allocation Funds, 7: Equity-oriented Hybrid Funds, 8: Debt-oriented Hybrid Funds, 9: Balanced Fund, 10: Equity Fund, 11: Enhanced Index Funds, 12: Enhanced Index Debt Funds, 13: Debt Funds(Medium- and long-term)
             
Deliverables：
- The generated document is automatically named after **"Market + Investment Type + Position Viewing Method"** and saved in the same path as the original file.

**Notes:** The program does not set a strict error reporting link, please make sure that the input is correct. If the input is incorrect but a new table is generated, the data is most likely inaccurate.
For example, if you want to choose REITs(1), Passive Funds(2), Stocks Long/Short(3), you should enter **"1 2 3"** in the input. If it loses to **"12 3"**, then the output data becomes Stocks Long/Short(3) and enhanced index bond funds(12). If the output is **"1 23"** then the output data becomes only REITs(1).
