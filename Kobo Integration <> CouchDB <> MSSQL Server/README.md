# Data Engineering: KoboToolbox → CouchDB → Microsoft SQL Server

A Python-based data engineering pipeline for extracting survey submissions from **KoboToolbox**, storing them in **CouchDB**, and synchronizing the data into **Microsoft SQL Server** for downstream analytics, reporting, and data integration.

The pipeline is designed for reliable, repeatable data ingestion and synchronization across multiple deployed KoboToolbox forms. It supports batch processing, incremental synchronization, database upserts, error handling, and continuous watch mode.

## Architecture

```text
┌─────────────────────┐
│    KoboToolbox      │
│                     │
│  Survey Forms       │
│  Submissions        │
└──────────┬──────────┘
           │
           │ REST API
           ▼
┌─────────────────────┐
│       Python        │
│    Data Pipeline    │
│                     │
│ Extract             │
│ Transform           │
│ Load / Sync         │
└──────────┬──────────┘
           │
           ▼
┌─────────────────────┐
│       CouchDB       │
│                     │
│ One database/form   │
│ Document storage    │
└──────────┬──────────┘
           │
           │ Changes API
           ▼
┌─────────────────────┐
│   Microsoft SQL     │
│      Server         │
│                     │
│ Analytics-ready     │
│ relational storage  │
└─────────────────────┘
```

## What the Pipeline Does

The pipeline automates the movement of survey data through three main systems:

### 1. Extract from KoboToolbox

The pipeline connects to the KoboToolbox API and:

* Retrieves deployed survey forms.
* Supports processing all deployed forms or a specific form.
* Retrieves survey submissions using the KoboToolbox API.
* Handles paginated API responses.
* Processes submissions in configurable batches.

Draft and undeployed forms are skipped automatically.

### 2. Load into CouchDB

Each KoboToolbox form is mapped to a dedicated CouchDB database.

The pipeline:

* Generates safe database names from form names.
* Adds the KoboToolbox form UID to each document.
* Preserves the submission as a JSON document.
* Performs bulk upserts.
* Detects unchanged documents and avoids unnecessary writes.
* Handles CouchDB document revisions.
* Reports failed document writes without stopping the entire pipeline.

### 3. Synchronize with Microsoft SQL Server

The pipeline then synchronizes CouchDB documents into Microsoft SQL Server.

For each CouchDB database, it:

* Creates the target SQL Server table when required.
* Tracks the last CouchDB change sequence.
* Uses the CouchDB `_changes` endpoint for incremental synchronization.
* Inserts new records.
* Updates changed records using SQL `MERGE`.
* Stores the complete source document as JSON in an `NVARCHAR(MAX)` column.
* Records import and update timestamps.
* Maintains synchronization state in `_couch_sync_state`.

This makes the pipeline suitable for repeated execution without unnecessarily reprocessing unchanged records.

## Key Data Engineering Features

### Incremental Synchronization

Instead of repeatedly loading the entire dataset into SQL Server, the pipeline maintains the last processed CouchDB sequence in:

```text
_couch_sync_state
```

The next synchronization starts from the previously saved sequence.

This enables incremental data ingestion and reduces unnecessary processing.

### Idempotent Upserts

Records are identified using the CouchDB document ID:

```text
couch_id
```

SQL Server uses `MERGE` to update existing records or insert new records.

This allows the pipeline to be safely executed repeatedly without creating duplicate records.

### Batch Processing

Data is processed in configurable batches to improve scalability:

```text
BATCH_SIZE=500
```

The same batching approach is used when writing documents to CouchDB and processing changes from CouchDB.

### Error Isolation

A failure involving one form does not necessarily stop processing of other forms.

The pipeline catches errors at the form and database synchronization levels and records failures through Python logging.

### Continuous Watch Mode

The pipeline can operate continuously using:

```bash
python pipeline.py --watch
```

The polling interval is configurable through:

```text
WATCH_POLL_SECONDS
```

This allows the pipeline to periodically check for new or changed survey submissions.

## Project Structure

A recommended repository structure is:

```text
data-engineering/
│
├── pipeline.py
├── .env
├── .gitignore
├── requirements.txt
├── README.md
│
├── data/
│   ├── raw/
│   └── processed/
│
├── sql/
│   └── ...
│
├── notebooks/
│   └── ...
│
└── docs/
    └── ...
```

> `.env` should contain credentials and connection information and **must not be committed to GitHub**.

## Technologies

| Technology             | Purpose                                        |
| ---------------------- | ---------------------------------------------- |
| Python                 | Pipeline orchestration and data processing     |
| KoboToolbox API        | Survey form and submission source              |
| CouchDB                | JSON/document-oriented intermediate data store |
| Microsoft SQL Server   | Structured downstream data storage             |
| `requests`             | HTTP/API communication                         |
| `pyodbc`               | SQL Server connectivity                        |
| `python-dotenv`        | Environment variable management                |
| SQL `MERGE`            | Insert/update synchronization                  |
| CouchDB `_changes` API | Incremental change detection                   |

## Requirements

The pipeline requires:

* Python 3.10+
* KoboToolbox account/API access
* CouchDB instance
* Microsoft SQL Server
* An appropriate SQL Server ODBC driver
* Network access between the pipeline and the required services

Install the Python dependencies with:

```bash
pip install -r requirements.txt
```

A typical `requirements.txt` for this project contains:

```text
pyodbc
requests
python-dotenv
```

## Configuration

Create a `.env` file in the same directory as the pipeline.

Example:

```env
# KoboToolbox
KOBO_BASE_URL=https://kf.kobotoolbox.org
KOBO_TOKEN=your_kobo_api_token

# CouchDB
COUCH_URL=http://localhost:5984
COUCH_USER=your_couchdb_user
COUCH_PASS=your_couchdb_password

# Microsoft SQL Server
DB_DRIVER=ODBC Driver 18 for SQL Server
DB_HOST=localhost
DB_PORT=1433
DB_NAME=your_database
DB_USER=your_username
DB_PASSWORD=your_password

# Pipeline configuration
BATCH_SIZE=500
WATCH_POLL_SECONDS=300
FAST_EXECUTEMANY=false
```

Alternatively, the pipeline supports providing a complete SQL Server ODBC connection string through:

```env
MSSQL_CONN_STR=your_connection_string
```

## Running the Pipeline

### Process all deployed KoboToolbox forms

```bash
python pipeline.py
```

The pipeline will:

1. Connect to KoboToolbox.
2. Discover deployed survey forms.
3. Retrieve submissions.
4. Load submissions into CouchDB.
5. Synchronize CouchDB data into SQL Server.
6. Close the SQL Server connection.

### Process a single form

Use the KoboToolbox asset UID:

```bash
python pipeline.py --form-id YOUR_FORM_UID
```

This is useful when testing a form or running a targeted data ingestion process.

### Synchronize existing CouchDB databases

If KoboToolbox extraction is not required:

```bash
python pipeline.py --sync-only
```

This mode skips KoboToolbox and synchronizes existing CouchDB databases into SQL Server.

### Continuous synchronization

To periodically repeat the pipeline:

```bash
python pipeline.py --watch
```

The polling interval is controlled by:

```env
WATCH_POLL_SECONDS=300
```

The default is 300 seconds (5 minutes).

## SQL Server Data Model

For each synchronized CouchDB database, the pipeline creates a corresponding SQL Server table containing:

| Column        | Description                        |
| ------------- | ---------------------------------- |
| `couch_id`    | Unique CouchDB document identifier |
| `couch_rev`   | CouchDB document revision          |
| `form_uid`    | KoboToolbox form UID               |
| `data`        | Complete submission stored as JSON |
| `imported_at` | Initial import timestamp           |
| `updated_at`  | Last synchronization timestamp     |

The pipeline also maintains:

```text
_couch_sync_state
```

This table stores the latest CouchDB change sequence processed for each database.

## Data Flow

The overall data flow can be summarized as:

```text
KoboToolbox Survey
        │
        ▼
KoboToolbox REST API
        │
        │ Extract
        ▼
Python Pipeline
        │
        │ Transform
        │ - Normalize identifiers
        │ - Add form metadata
        │ - Prepare JSON documents
        ▼
CouchDB
        │
        │ Incremental _changes feed
        ▼
Python Pipeline
        │
        │ Upsert
        ▼
Microsoft SQL Server
        │
        ▼
Analytics / Reporting / BI
```

## Logging and Monitoring

The pipeline uses Python's built-in logging framework.

Logs include information such as:

* Number of KoboToolbox forms discovered.
* Number of deployed forms.
* Number of submissions retrieved.
* CouchDB documents written.
* Unchanged documents skipped.
* Document-level errors.
* SQL Server synchronization status.
* Number of synchronized documents.
* Synchronization failures.
* Watch-mode polling cycles.

Example log messages:

```text
KoboToolbox: found X form(s), Y deployed
CouchDB: X written, Y unchanged, Z errors
MS SQL Server: synced X doc(s)
MS SQL Server: 'table_name' already up to date
```

## Design Principles

This project follows several practical data engineering principles:

**Reliability**
Failures are isolated where possible so that one problematic form or database does not stop the entire pipeline.

**Idempotency**
Repeated executions do not create duplicate records.

**Incremental Processing**
CouchDB change sequences are used to avoid repeatedly processing unchanged data.

**Scalability**
Batch processing and optional `fast_executemany` support are used to improve performance when handling larger datasets.

**Traceability**
Source identifiers, form UIDs, document revisions, and timestamps are retained.

**Separation of Concerns**
The pipeline separates extraction from KoboToolbox, document storage in CouchDB, and relational synchronization into SQL Server.

## Performance Considerations

The pipeline provides configurable batch processing through:

```env
BATCH_SIZE=500
```

For supported Microsoft ODBC drivers, `fast_executemany` can be enabled:

```env
FAST_EXECUTEMANY=true
```

The implementation also falls back to normal execution if fast execution encounters an ODBC error.

For larger deployments, performance can be further improved by tuning:

* Batch size
* SQL Server indexes
* ODBC configuration
* CouchDB database configuration
* Network connectivity
* SQL Server hardware/resources

## Security

Credentials are loaded from environment variables using `python-dotenv`.

Do not commit credentials to source control.

Add the following to `.gitignore`:

```gitignore
.env
__pycache__/
*.pyc
.venv/
venv/
```

API tokens, database passwords, and connection strings should be managed through environment variables or a secure secrets-management solution.

## Use Cases

This pipeline can support:

* Survey data ingestion
* Field data collection systems
* KoboToolbox data warehousing
* Monitoring and evaluation data pipelines
* Research data infrastructure
* Operational reporting
* Business intelligence
* Centralized SQL Server data stores
* Automated data synchronization
* Near-real-time or scheduled survey data ingestion

## Future Enhancements

Potential extensions include:

* Automated schema inference from KoboToolbox forms
* Relational normalization of nested survey data
* Data quality validation rules
* Pipeline orchestration with Airflow or similar tools
* Automated data quality monitoring
* SQL Server dimensional modeling
* Power BI integration
* Structured audit logging
* Retry and backoff strategies
* Containerization with Docker
* CI/CD deployment
* Automated testing
* Data lineage and metadata management

## Project Status

**Status:** Active development

This repository is intended as a practical data engineering implementation demonstrating API-based extraction, document-oriented storage, incremental synchronization, and relational database integration.

## License

Add the appropriate license for your project here.

For example:

```text
MIT License
```

if the project is intended to be released under the MIT License.

