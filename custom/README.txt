WFMHub Custom Lab
=================

Python
------
Copy jobs\_paste_your_python_here.py, remove the leading underscore from the
new filename, and paste custom code inside run(ctx). The hub connection exposed
through ctx is read-only. Python jobs are trusted local code: only run code you
understand. The portable runtime supports the Python standard library plus the
Excel libraries shipped with WFMHub; it does not run pip on the work computer.

SQL
---
Copy sql\_paste_your_sql_here.sql, remove the leading underscore, and edit the
single SELECT/WITH query. Use :start and :end for the dates chosen in the menu.

Outputs are written under output\custom and source extracts are not changed.
