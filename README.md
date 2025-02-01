# ExcelDBSeeder

## Overview
Excel to SQL Importer is a C# application that reads large Excel tables and inserts the data into normalized relational database tables. It simplifies data migration by automatically parsing and distributing data into appropriate SQL tables.

## Features
- Reads large Excel files efficiently.
- Parses and normalizes data before inserting it into relational tables.
- Supports bulk insert operations for improved performance.
- Handles missing or incorrect data gracefully.
- Logs errors and successful operations for debugging.

## Prerequisites
Ensure you have the following installed before running the application:

- .NET 6.0 or later
- Microsoft SQL Server
- ADO.NET
- EPPlus (for Excel processing)

## Installation

1. Clone the repository:
   ```sh
   git clone https://github.com/your-username/ExcelToSQLImporter.git
   cd ExcelToSQLImporter
