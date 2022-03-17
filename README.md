# VBA-ETL

A framework for bulk loading data from any ADO-compliant data source to SQL Server (or, theoretically, any database, though I've only coded the classes for SQL Server and some generic ANSI types).

Data is loaded asynchronously in bulk in an effort to minimize waits and network traffic.

Possible enhancements:
- Load varchar/nvarchar fields as binary to completely eliminate SQL injection risk (theoretically you're loading from a safe data source, and there's some effort to properly escape character strings, but that's the more paranoid way to do it).
- Update data loading classes to support loading to existing named tables (rather than only temp tables).
- Classes to simplify indexing temp tables.
- A wrapper to simplify creating and loading tables.
- Option to submit data as a single GZipped parameter that is decompressed and executed server side (testing showed some performance improvements for large data sets over a VPN, and ~70% compression).
