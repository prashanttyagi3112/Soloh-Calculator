# Soloh Travel Pay Package Calculator v1.1 (18-18)

Excel calculator for Soloh Partners recruiters to check margin on travel nurse pay packages before making an offer.

## Rules built in
| Setting | Value |
|---|---|
| Payroll load | **18% on W2 only** (RT + OT + W2 bonuses; not on per diem) |
| Onboarding cost | **$0** |
| Desired GP / net margin | **18%** of revenue |
| Orientation | Non-billable orientation hours are paid to the nurse but not billed (default 12 hrs) |
| Client OT | "Client bills OT? = No" bills every hour at straight time (e.g. Norton up to 48 hrs) |

All settings sit in the **Admin Use Only** box and can be changed.

## Use it online
**Web calculator:** `https://prashanttyagi3112.github.io/Soloh-Calculator/` (works on phone and laptop – no Excel or macros needed; includes Clear/Reset, Copy Offer Summary and Print/Save PDF).

Nothing typed into the web calculator is saved or sent anywhere – all math runs in the browser.

## How to use (Excel version)
1. Download `Soloh_Travel_Pay_Package_Calculator_v1.1_18Load-18GP.xlsx` and open it in Excel (no macros needed).
2. Fill the **yellow cells**: client, MSP %, bill rate, W2 rate, OT rate, weekly per diem, hours/week, weeks, orientation hours, bonuses.
3. Read **Net Margin %** and the green/red check against 18%.
4. If it's red, use **Offer Guidance** for the max per diem or max W2 rate that still hits 18%.
5. Share the **Pay Package Confirmation** panel (File → Save as PDF, one page).

Formula cells are locked (password: ask Prashant).

## Optional buttons (macros)
`SolohCalc_Macros.bas` adds **Clear / Reset** and **Save Offer PDF** buttons. Setup steps are on the **Macro Setup** tab inside the workbook. Save as `.xlsm` after adding.

## Versions
- **v1.0** – 18% load on W2, $0 onboarding, 18% desired GP.
- **v1.1** – Removed Candidate Details, Job Title, JobDiva Ref # and Client Job #; the five bonus lines are now one Bonus / Reimbursement field (W2, loaded at 18%); Blended Hourly Pay Package section removed.
- **v1.1 update** – Payroll Load (18%) is locked. New Admin option "Target Net Margin $": enter a dollar amount and the W2 RT rate is set automatically to hit it (everything else stays the same); leave blank / 0 to type W2 manually.
