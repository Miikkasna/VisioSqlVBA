# VisioSql - SQL from Visio Database Diagrams

VBA script to forward engineer Visio Database Diagram to SQL-scripts. Tested on Visio Professional 2021.

The script iterates through all pages in the Visio document and all shapes in pages and extracts information from table-shapes. "CREATE TABLE"-scripts are saved to separate sql-files per page.

Primary Keys and Foreign Keys must be adapted to specific SQL-language (currently using postgreSQL format). Foreign Key references are assumed using field-names.
