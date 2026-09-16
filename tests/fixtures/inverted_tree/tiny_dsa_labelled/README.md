# Tiny DSA with runtime axis labels

Copy of `../tiny_dsa` whose `TIME_PERIOD` headers are formulas of
`first_projection_year` (`Inputs!B3`). Bind `engine_year_labels` with
`axis_labels: TIME_PERIOD` so generated packages key every tensor on that
axis by the evaluated labels.

The literal-header workbook in `../tiny_dsa` stays the unlabelled canary.
Other header rows (`Engine!C13:G13`, `Outputs!B11:F11`) keep snapshot
integers so catalog key resolution does not require those cells in the
graph; they must match the labeller's snapshot values.
