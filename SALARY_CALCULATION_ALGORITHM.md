# Salary Calculation Algorithm

This document describes how the salary slip app calculates a monthly salary from the employee Excel data.

## Excel entry rule

- **Present Days** includes Sundays and every other calendar day the employee is not on leave.
- Do not subtract Sundays separately. They are already part of Present Days.
- **Paid Leave** contains only paid leave days. Enter it separately.
- Unpaid leave is inferred by the app; there is no separate unpaid-leave input.
- Present Days + Paid Leave should not exceed the number of calendar days in the selected month.

## Formula

Let:

- `D` = calendar days in the selected month (28, 29, 30, or 31)
- `P` = Present Days from Excel, with Sundays included
- `L` = Paid Leave from Excel
- `G` = Basic Salary + Allowance
- `F` = Other / Fixed Deduction from Excel

Calculate:

```text
Paid Days                 = min(P + L, D)
Unpaid Leave Days         = max(D - Paid Days, 0)
Salary Per Calendar Day   = G / D
Unpaid Leave Deduction    = Salary Per Calendar Day * Unpaid Leave Days
Total Deductions          = Unpaid Leave Deduction + F
Net Salary                = max(G - Total Deductions, 0)
```

Amounts are rounded to two decimal places for the slip. Sundays are not added again as a separate attendance value: they are already included in `P` and in the calendar-day divisor `D`.

## Worked example

For a 31-day month, with 5 leave days (2 paid and 3 unpaid):

```text
Present Days (Excel) = 31 - 5 = 26
Paid Leave (Excel)   = 2
Paid Days            = 26 + 2 = 28
Unpaid Leave Days    = 31 - 28 = 3
```

If Basic Salary is ₹30,000, Allowance is ₹0, and fixed deductions are ₹0:

```text
Salary Per Calendar Day = ₹30,000 / 31 = ₹967.74
Unpaid Leave Deduction  = ₹967.74 * 3 = ₹2,903.23
Net Salary              = ₹30,000 - ₹2,903.23 = ₹27,096.77
```

This gives credit for the 2 paid leave days and deducts pay for the 3 unpaid leave days. Sundays remain included in the 31-day month calculation.
