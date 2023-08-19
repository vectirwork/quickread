# quickread
Tool that automate memory caching of file used by other modules.
For example if you are using openpyxl to read-only Excel files, consider using openpyxl.open=quickread.Define(openpyxl.open).open
If you are inserting this line at the top of an existing script, it can speed up your python code!
