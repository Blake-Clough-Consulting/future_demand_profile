# future_demand_profile

Basic way of shifting and scaling the metered GP9 data using regional FES to come up with a future profile.

The raw `Meter Volume` column is preserved. Profile scaling uses `Signed Meter Volume`, where GP9 rows marked with `Import/Export Indicator = E` are treated as negative power flow.

When multiple settlement runs exist for the same GSP and half-hour, the profile keeps the best available run using the priority order `R3 > R2 > R1 > SF > RF > DF > II` and prints a summary of the selected run types.
