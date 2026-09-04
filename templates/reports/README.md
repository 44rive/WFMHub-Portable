# Excel master policy

The only persistent Excel master is `Reports/PCS Team.xlsx`.

WFMHub creates that data-free starter during portable packaging or from the
`pcs-team` command when it is missing. Daily refreshes update only the compact
CSV feeds under `_system/feeds/pcs/current`; they do not overwrite the team's
PivotTables, slicers, or coaching log.

`SETUP.cmd` renders firewall-safe, installation-specific query scripts into
`_system/power_query` for the one-time Excel setup.

All other workbooks are replaceable snapshots published directly in `Reports`
with their previous copies stored under `Reports/Archive`.
