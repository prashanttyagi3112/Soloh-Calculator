# Soloh Travel Pay Package Calculator v1.0 (18-18)

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

## How to use
1. Download `Soloh_Travel_Pay_Package_Calculator_v1.0_18Load-18GP.xlsx` and open it in Excel (no macros needed).
2. Fill the **yellow cells**: client, MSP %, bill rate, W2 rate, OT rate, weekly per diem, hours/week, weeks, orientation hours, bonuses.
3. Read **Net Margin %** and the green/red check against 18%.
4. If it's red, use **Offer Guidance** for the max per diem or max W2 rate that still hits 18%.
5. Share the **Pay Package Confirmation** panel (File → Save as PDF, one page).

Formula cells are locked (password: ask Prashant).

## Optional buttons (macros)
`SolohCalc_Macros.bas` adds **Clear / Reset** and **Save Offer PDF** buttons. Setup steps are on the **Macro Setup** tab inside the workbook. Save as `.xlsm` after adding.

## Versions
- **v1.0** – 18% load on W2, $0 onboarding, 18% desired GP.
