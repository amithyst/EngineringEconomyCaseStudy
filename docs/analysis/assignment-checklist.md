# Assignment Checklist

## Required Outputs

- annual cash flow model for 2026-2029
- acquisition price recommendation under `MARR = 12%`
- sensitivity analysis on leasing assumptions
- open-ended discussion on non-financial risks and tenant strategy

## Cash Flow Items To Include

- acquisition price at 2026-01-01
- deed tax on acquisition price
- other transaction fees
- mall opening cost
- rental income
- parking income
- property management fee income
- operating costs
- corporate income tax
- exit value at 2029-12-31

## Timeline Constraints

- acquisition date: `2026-01-01`
- reopening date: `2026-04-01`
- first-year rent-free period: `3 months`
- first-year rent collected: `6 months`
- first-year operating period for cost allocation: `9 months`
- exit date: `2029-12-31`

## Explicit Modeling Rules From The Prompt

- ignore depreciation
- property fees are still charged during rent-free months
- parking revenue should depend on actual operating days
- operating costs should be scaled by actual operating time in 2026
- terminal value should be based on a perpetuity with `i = 8%`
- corporate income tax rate is fixed at `25%`
- terminal value is not taxed

## Questions The Submission Should Answer

- how quickly does cash flow improve from the renovation year to the stable years
- how different are the renovation-period and stabilized-period economics
- how sensitive is valuation to occupancy and rent growth
- what strategic or operating risks are not captured by the financial model
