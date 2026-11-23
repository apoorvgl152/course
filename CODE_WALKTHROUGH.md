# How the Glue ingestion job works

This is the reading guide for the metadata-driven Glue script (the module whose header starts with `METADATA-DRIVEN GLUE INGESTION`). The script is the whole pipeline: config in, Parquet in S3 out. Nothing about a particular source is hard-coded.

Read it in the order below. Each step names the functions to open. Stop at the end of a step and match it to the code before going on.

## How to proceed

1. Read [What one run is](#what-one-run-is) and [The four tables](#the-four-tables). That is the whole shape.
2. Read [From `main()` to the first object](#from-main-to-the-first-object) with the script open. Follow the calls. Do not read the helpers yet.
3. Read [One object, extract to Parquet](#one-object-extract-to-parquet) the same way. This is the part that moves data.
4. Read [Incremental watermarks](#incremental-watermarks) next. Most of the subtle behavior is there.
5. Use [When something fails](#when-something-fails) and [Map of the file](#map-of-the-file) as lookup. Open a function only when a step names it.

The file is organized as ten numbered sections (logging, constants, retry, I/O, audit, metadata, config, processing, orchestration, entry point). Runtime order is the reverse of that: section 10 calls section 9, which calls section 8, which calls section 4. The sections exist so the helpers sit still while you read the drivers.

## What one run is

A Glue run is started by Step Functions with `--SOURCE_NAME`, for example `FISERV` or `LMS,FISERV`. The script loads every active `JOB_CONFIG` row whose `JOB_TYPE` is `GLUE` and whose `SOURCE_NAME` is in that list. Each of those rows is one job. Each active `OBJECT_CONFIG` row under the job is one object (one table, one query, or one S3 prefix).

```text
Glue run
  job  (one JOB_CONFIG row, its own RUN_ID)
    object  (one OBJECT_CONFIG row)
      read source  ->  Parquet in S3  ->  OBJECT_LOG
    object ...
  job ...
```

Jobs run at the same time, up to 3. Objects inside a job run at the same time, up to 4. A process-wide semaphore caps extracts at 12 so the two limits cannot stack into 12-per-job. Audit writes do not take the semaphore.

Two clocks are fixed when the module is imported, in UTC, and reused for every object in the process:

| Name | Example | Used for |
| --- | --- | --- |
| `LOAD_DATE` | `2026-08-25` | The date folder under the target prefix |
| `RUN_TIMESTAMP` | `20260814_155052` | The stamp inside the file name |

A run that crosses midnight stays in the folder it started in. A second Glue run gets a new stamp and writes a new file beside the old one. Earlier files are never deleted.

## The four tables

All four live in `EDH_POC_DB.METADATA_CONFIG`. Onboarding a source is an insert into the first two.

| Table | Grain | Role |
| --- | --- | --- |
| `JOB_CONFIG` | one row per job | Which system, which Glue connection, where in S3 the files go, who to email |
| `OBJECT_CONFIG` | one row per object | The query or the file prefix, load type, watermark, file name, write mode |
| `JOB_LOG` | one row per job run | `STARTED`, then updated to `SUCCESS` or `FAILED` |
| `OBJECT_LOG` | one row per object run | Same lifecycle, plus row counts and `LAST_WATERMARK_VALUE` |

`JOB_LOG` and `OBJECT_LOG` are Snowflake hybrid tables. Their ids come from `JOB_LOG_SEQ` and `OBJECT_LOG_SEQ`. `OBJECT_LOG.JOB_LOG_ID` points at the parent job row.

Columns the script actually branches on:

| Column | Table | What the script does with it |
| --- | --- | --- |
| `SOURCE_DB` | job | `SNOWFLAKE`, `ORACLE`, `SQLSERVER`, or `S3`. Anything else fails the object. |
| `SOURCE_NAME` | job | Matched to `--SOURCE_NAME` with `UPPER(TRIM())` on both sides. Also the first folder of the target path, compared in lower case. |
| `CONNECTION_INFO` | job | Glue connection name for Oracle and SQL Server. Snowflake uses `--SNOWFLAKE_CONNECTION` instead. |
| `S3_TARGET_FILE_PATH` | job | Bucket root. Must start with `s3://`. |
| `SOURCE_FILE_FORMAT` | job | `csv` or `dat` for S3 sources. |
| `SOURCE_FILE_DELIMITER` | job | Required for S3 sources. Passed through as stored. |
| `JOB_FAILURE_NOTIFICATION` | job | Comma-separated email addresses. Used only after a failure. |
| `SOURCE_QUERY` | object | The SELECT for a database source. Wrapped, not edited, when a watermark applies. |
| `SOURCE_S3_PATH` | object | Prefix of the source files for an S3 source. |
| `SOURCE_SCHEMA` | object | Optional JSON column list for S3 reads. |
| `DATA_LOAD_TYPE` | object | The incremental path runs only when this is exactly `INCREMENTAL`. |
| `WATERMARK_FIELD` | object | Required when the load type is `INCREMENTAL`. Compared with `>`. |
| `WRITE_MODE` | object | `overwrite` or `append`. Required. See [the write](#the-write). |
| `TARGET_OBJECT_FILE_FORMAT` | object | Must be `parquet` by the time a non-empty extract is written. |
| `TARGET_FILENAME_CONVENTION` | object | Stem of the output file, for example `ACCTCALCVAR_FISERV.PARQUET`. |
| `TARGET_S3_FILE_PATH` | object | Exactly four folders: `source/ingestion_type/frequency/object_name`. |
| `PARTITION_KEY` | object | Optional. A date or timestamp is rewritten to `<column>_YEAR`. |
| `PARTITION_SIZE_IN_MB` | object | Snowflake connector only. Set on a JDBC source, the object fails. |
| `PARTITION_NO_OF_ROWS` | object | Optional cap on rows per output file. |

`IS_ACTIVE` must be true on the job and on the object. `JOB_TYPE` must be `GLUE`.

## From `main()` to the first object

Open the script at the bottom and walk up.

`main` records a start time, calls `start_run`, and logs one line: `Glue run SUCCESS` or `Glue run FAILED`. The failure line includes the traceback. The exception is raised again so Glue marks the run failed.

`start_run`:

1. `resolve_job_arguments` reads `JOB_NAME`, `SES_FROM_ADDRESS`, `SNOWFLAKE_CONNECTION`, and `SOURCE_NAME`. A missing argument is Glue's error. A present but blank one is rejected here. `SOURCE_NAME` is split on commas, trimmed, and de-duplicated.
2. Those values are stored on `RUNTIME`. Later code calls `RUNTIME.require(...)` so a missing sender or Snowflake connection fails with its name, instead of a boto or connector error about `None`.
3. A `SparkContext` is created with `spark.scheduler.mode=FAIR`, then `GlueContext` and `Job.init`. FAIR scheduling is what lets the parallel jobs share the cluster.
4. `run_all_jobs` is the rest of the run.

`run_all_jobs` makes one id (`glue_run_id`) for the orchestration log and for a run-level failure email. Each job then gets a different id from `new_run_id()`. A job id looks like `20260821172114448_<uuid>`: UTC time to the millisecond, then a UUID, so ids sort and still differ when two jobs start in the same millisecond.

`run_jobs_concurrently`:

1. `check_sequence_drift` for both log tables. A warning when the sequence sits at or behind `MAX(id)`. It does not stop the run. Inserts retry on a real collision.
2. `get_active_jobs`. No matching job, or two active rows with the same `JOB_ID`, raises. That raise is a run-level failure: no single job owns it.
3. `get_active_objects_bulk`. One query for every job. A job with no objects is stored as an empty list and fails later inside that job, not here.
4. The main thread closes its Snowflake connection. The job threads open their own.
5. A pool of up to 3 threads runs `run_pipeline` once per job.
6. `glue_job.commit()` runs after the pool finishes, including when some jobs failed.
7. The list of `(JOB_ID, error)` comes back. `run_all_jobs` raises one `RuntimeError` that names each failed job. Each of those jobs has already sent its own email.

`run_pipeline` is one job:

1. Insert `JOB_LOG` as `STARTED`. This insert is allowed to raise. Without a `JOB_LOG_ID`, object rows would have no parent.
2. Zero active objects is a failed job, not a quiet success.
3. `run_job_objects` runs the objects.
4. Any object with status `FAILED` fails the job. The job message lists the object ids. The per-object reason is already on `OBJECT_LOG` and in the email.
5. Otherwise the job row is updated to `SUCCESS`.

The job's `except` block updates `JOB_LOG` to `FAILED`, calls `notify_job_failure`, and re-raises. The update and the email are guarded: a Snowflake or SES error there is logged and does not replace the original exception.

## One object, extract to Parquet

`run_job_objects` loads watermarks for the job's `INCREMENTAL` objects in one query (`get_last_watermarks_bulk`), then submits every object to a pool of up to 4 threads. Results are put back in config order. When the pool exits, Snowflake connections whose thread name starts with `obj-<JOB_ID>_` are closed.

`run_object` is the boundary that catches errors. One bad object does not cancel the others.

1. Insert `OBJECT_LOG` as `STARTED`. If that insert fails, the object still runs. The final status is inserted at the end instead of updated.
2. `build_target_s3_path` builds the destination. The same string is what gets logged.
3. Under `global_object_semaphore`, `process_object` reads and writes.
4. On success, update the log row with counts and, for an incremental object, the new watermark.
5. On any exception, update the log row to `FAILED` with `str(exc)` and return a result dict. The exception does not leave `run_object`.

`process_object` is the extract and the load. The comments in the function number the steps.

**Step 1, before any read.** `check_object_preflight` rejects a null `WRITE_MODE`, and a null `WATERMARK_FIELD` when the load type is `INCREMENTAL`. `build_target_file_name` requires `TARGET_FILENAME_CONVENTION`.

**Step 2, the resume point.** For an incremental object the watermark argument is the value from `OBJECT_LOG`. For anything else it is ignored. The first incremental run has no prior success, so the watermark is null and the extract is unfiltered. That run is marked `is_first_incremental` so the log message says so.

**Step 3, a lazy frame.** `extract_source` switches on `JOB_CONFIG.SOURCE_DB`. Nothing has been read yet, except where the connector itself resolves a schema (JDBC `load()`, and an S3 read that infers types).

**Step 4, partition column.** If `PARTITION_KEY` is set, `derive_partition_column` looks at the field type. A date or timestamp gains `<column>_YEAR`, and the write partitions by the year. Any other type is used as-is. Strings are not parsed as dates. A missing column, or a year column that already exists, fails the object.

**Step 5, one materialization.** The frame is cached. `count()` is the action that actually reads the source. The cache is then reused for the write and for `MAX(watermark)`, and `unpersist()` runs on the way out.

- Count is 0: nothing is written. An incremental object keeps the watermark it started with.
- Count is greater than 0: an S3 `FULL` load may run the raw-line check, then `write_to_s3_with_named_file`. The new watermark is `MAX(WATERMARK_FIELD)` over the cached rows, which is exactly the set that was written.

The return value is `row_count`, `latest_watermark`, and `is_first_incremental`.

If the data was written and the `OBJECT_LOG` success update did not land, an `INCREMENTAL` object raises. The next run would otherwise look like a first load and extract the same rows again. A `FULL` object in that situation still returns success. The data is already in S3, and a full load has no watermark to lose.

### The path

`build_target_s3_path` joins three pieces:

```text
JOB_CONFIG.S3_TARGET_FILE_PATH          s3://bucket/landing/edh
OBJECT_CONFIG.TARGET_S3_FILE_PATH       fiserv/batch/daily/acctcalcvar
LOAD_DATE                               2026-08-25
= s3://bucket/landing/edh/fiserv/batch/daily/acctcalcvar/2026-08-25/
```

`parse_target_s3_file_path` checks the middle piece before any I/O, because S3 would accept a wrong prefix and the object would still look successful. The value must be four non-empty components:

```text
<source_name>/<ingestion_type>/<frequency>/<object_name>
```

`source_name` is `JOB_CONFIG.SOURCE_NAME` in lower case. `ingestion_type` is `batch` or `streaming`. `frequency` is `daily` or `monthly`. One leading slash and one trailing slash are stripped. A doubled slash is reported, not repaired.

The file name inserts `RUN_TIMESTAMP` before the extension, and the extension is upper-cased:

```text
ACCTCALCVAR_FISERV.PARQUET
  -> ACCTCALCVAR_FISERV_20260814_155052.PARQUET
```

### The write

`write_to_s3_with_named_file` does not ask Spark to choose the final name.

1. Spark writes Parquet under `target_path/_tmp_<uuid>/`. Names that start with `_` are skipped by Spark, Athena, and Glue crawlers, so a folder left by a killed run is not read as data.
2. Part files (`part-*.parquet`) are grouped by partition folder. One part in a folder is copied to the convention name. Several parts become `stem_00001.ext`, `stem_00002.ext`, and so on.
3. Each copy is checked with `head_object`. The size must match the temp object.
4. The temp prefix is deleted on the way out, success or failure. A failed delete is logged and swallowed. It must not hide a write error, and it must not fail an object whose files are already published.

`WRITE_MODE` is passed to the Spark write of that temp folder. It does not delete older files in the date folder.

- `overwrite` replaces the temp folder, so a transient error during the Spark write is retried.
- `append` is not retried. A retry after a partial temp write could duplicate rows inside the scratch folder.

`PARTITION_NO_OF_ROWS`, when set, does two things together: `repartition` sets how many tasks write, and `maxRecordsPerFile` caps rows in a single file. Partitioned writes repartition by the partition column so one value stays in one task.

## Two ways the script talks to Snowflake

Bulk object data goes through Spark. Config, watermarks, sequence values, and audit rows go through the Snowflake Python connector on the calling thread. The Spark connector's staging writes and `NEXTVAL` are not safe when many threads hit the same log table.

Credentials follow one chain, fetched once per connection name per run and reused: Glue connection, then `AuthenticationConfiguration.SecretArn`, then Secrets Manager. The secret body is not logged.

| Path | Functions | Connection |
| --- | --- | --- |
| Snowflake bulk read | `read_snowflake` | Glue connection named by `--SNOWFLAKE_CONNECTION`, via `create_data_frame.from_options` |
| Oracle / SQL Server bulk read | `read_jdbc` | Glue connection named by `JOB_CONFIG.CONNECTION_INFO`. URL is built from the secret's host, port, and database. |
| S3 bulk read | `read_s3` | The job's IAM role. No database secret. |
| Control plane | `dbapi_cursor`, `dbapi_query_as_rows`, `execute_dbapi_write` | Per-thread `snowflake.connector` connection, key-pair auth from the same Snowflake secret (`sfUser`, `pem_private_key`) |

A thread's DB-API connection is reused. Every five minutes it is checked with `SELECT 1`. A transient error inside `dbapi_cursor` drops the connection so the retry opens a new one. `run_object` leaves the connection open for the thread's next object. The pool owner closes them after the pool drains.

JDBC secrets must contain `host`, `port`, `database`, `username`, and `password`. `fetchsize` is 10000. A trailing semicolon on `SOURCE_QUERY` is removed because Spark wraps the text in its own subquery.

## Incremental watermarks

The resume value is the newest `OBJECT_LOG.LAST_WATERMARK_VALUE` on a `SUCCESS` row for that `OBJECT_ID` (`QUALIFY ROW_NUMBER() ... ORDER BY OBJECT_LOG_ID DESC`). A failed run does not move it. An object with no such row is a first incremental load.

For a database source, `build_source_query` leaves `SOURCE_QUERY` unchanged on a full load and on a first incremental load. On a later incremental load it wraps the text:

```sql
SELECT * FROM (
  <SOURCE_QUERY with one trailing semicolon removed>
) src          -- " src" on Oracle and SQL Server, empty on Snowflake
WHERE <WATERMARK_FIELD> > <literal>
```

The comparison is strictly greater than, because the row at the previous watermark was loaded last time. The literal is rendered per engine:

| `SOURCE_DB` | Literal |
| --- | --- |
| `SNOWFLAKE` | `TO_TIMESTAMP_NTZ('YYYY-MM-DD HH:MM:SS.ffffff')` |
| `ORACLE` | `TO_TIMESTAMP(..., 'YYYY-MM-DD HH24:MI:SS.FF6')` |
| `SQLSERVER` | `CONVERT(DATETIME2(6), ..., 121)` |

Wrapping is what keeps the query's own `WHERE`, `GROUP BY`, CTEs, and line breaks intact. The watermark column has to be in the query's output. `get_latest_watermark` reads it from the cached frame after the extract.

For an S3 source there is no SQL to push the predicate into. The same `column > watermark` filter is applied on the DataFrame.

After a non-empty extract, the value stored for next time is `MAX(WATERMARK_FIELD)` of the rows just written. After an empty extract, the previous watermark is stored again.

Prefer a timestamp column. A date watermark only moves by whole days, so a later change on the same day is left behind.

## S3 sources

`read_s3` uses Spark's CSV reader for both `csv` and `dat`. `SOURCE_FILE_DELIMITER` is the separator. Every file is expected to have a header row.

`SOURCE_SCHEMA` is optional JSON:

```json
[
  {"name": "id", "type": "string"},
  {"name": "created_at", "type": "timestamp", "format": "yyyy-MM-dd HH:mm:ss"}
]
```

Types the parser accepts: `string`, `integer`, `long`, `float`, `date`, `decimal` (38, 10), `timestamp`. Spark allows one `dateFormat` and one `timestampFormat` per read, so the first `format` on a date column and the first on a timestamp column apply to the whole read. With a schema, types are not inferred. Without one, `inferSchema` is on.

For an S3 `FULL` load with at least one parsed row, `check_full_load_row_count` reads the prefix a second time as raw text, drops blank lines, subtracts one line per file for the header, and compares that to the parsed count. A mismatch fails the object. The check is there to catch a parser that dropped rows. It compares counts only.

## When something fails

Three scopes, one email each. A success sends nothing. Object failures do not send their own mail. They roll into the job email.

| Scope | When | Recipients |
| --- | --- | --- |
| Object | Any exception inside `run_object` | None. The job email includes the object. |
| Job | `run_pipeline` raises, including "no active objects" and "N objects failed" | That job's `JOB_FAILURE_NOTIFICATION` |
| Run | An exception outside any one job: Snowflake down during metadata, no active jobs, duplicate `JOB_ID` | Every notification address on the active jobs in `--SOURCE_NAME` |

The job email has the job identity, the timing, an object table (successes and failures), and a second table of the failed objects with `str(exc)`. The run email has no object table, because no job got that far. If the recipient lookup itself fails, the error stays in CloudWatch. The original exception is still raised.

`JOB_LOG` then stores whether the job email went out: `FAILURE_EMAIL_SENT_FLAG` `Y` or `N`, the recipient list, the time, and the SES error text.

`retry_call` is the only retry. It retries a call that does the I/O, not a call that only builds a lazy plan. Transient means a connection or timeout error, a Snowflake `OperationalError`, or a botocore `ClientError` whose code or HTTP status is throttling or 5xx. A missing credential, a bad parameter, and a config `ValueError` are not retried. The wait is capped exponential backoff plus jitter, four attempts by default.

Audit inserts have a second, narrower retry. The log tables enforce a primary key. If the insert hits "already exists" or "duplicate key":

- the row already stored under that id is this insert (same run, object, status, and microsecond timestamp) — the ack was lost after commit, so the insert is done;
- otherwise the sequence is behind `MAX(id)` — a new `NEXTVAL` is taken and the insert is tried again.

The new id is written into the row dict in place. The caller reads it back to learn which id landed.

Log lines are one JSON object on stdout: time, level, logger, thread, message, plus `run_id`, `JOB_ID`, and `object_id` from thread-local storage. Messages do not repeat those ids, except the two `as_completed` loops, which run on the parent thread and so put the child id in the text. Library loggers (`snowflake.connector`, boto, `py4j`, `py.warnings`) emit each distinct warning once per process. `ERROR` and the pipeline's own loggers are not filtered.

## Map of the file

| Section | What to open it for |
| --- | --- |
| 1. Logging | Why a line looks like JSON, and where `run_id` / `JOB_ID` / `object_id` are attached |
| 2. Constants | Table names, caps, the frozen date and timestamp, dialects, JDBC drivers, log column order |
| 3. Retry | The list of errors that are retried |
| 4. Source / S3 I/O | Connections, the DB-API helpers, `read_snowflake`, `read_jdbc`, `read_s3` |
| 5. Audit logging | The `STARTED` insert and the status update |
| 6. Metadata readers | The three bulk reads: jobs, objects, watermarks |
| 7. Config access | `get_config_value` and the path / schema parsers |
| 8. Processing | `process_object` and the path, name, query, and write helpers |
| 9. Orchestration | `run_object`, `run_pipeline`, `run_all_jobs`, and the emails |
| 10. Entry point | `main`, `start_run`, the four Glue arguments |

### Function index

Use this when you are standing on a name and want its job. The narrative above is the order to learn them.

**Logging.** `set_log_context` / `clear_log_context` store the ids for the current thread. `LogContextFilter` copies them onto every record. `JsonLogFormatter` prints one JSON line. `RepeatedLibraryRecordFilter` drops repeated library warnings. `configure_root_json_logging` installs that handler once. `utc_now` is naive UTC, matching `TIMESTAMP_NTZ`. `elapsed_seconds` is the duration every log, audit row, and email uses. `log_outcome` is the closing line of a run, job, or object. `new_run_id` builds the sortable id. `log_level_for` stores `ERROR` beside `FAILED` and `INFO` otherwise. `split_comma_list` parses `--SOURCE_NAME` and the notification list. `to_decimal` feeds `NUMBER` columns. `RuntimeConfig.require` fetches a Glue argument that `start_run` stored.

**Retry.** `is_transient_error` is the line between a blip and a real failure. `backoff_delay` is the wait. `retry_call` runs the call again from scratch, so the call has to be safe to repeat.

**I/O.** `sql_string_literal` / `sql_in_list` quote every `IN` value, including zero-padded ids like `0001`. `active_glue_jobs_predicate` is the `WHERE` for "active GLUE jobs in these sources". `split_s3_uri`, `list_s3_objects`, and `delete_s3_keys` are the S3 primitives. `load_private_key_der` turns the PEM into the bytes the connector wants. `get_glue_connection` and `read_secret` are the Glue-to-Secrets step. `get_snowflake_connection_properties` and `get_jdbc_connection_properties` build the two property dicts. `shared_connection_properties` caches them for the process. `get_snowflake_dbapi_connection` / `close_thread_snowflake_connection` / `close_snowflake_connections` own the per-thread connections. `dbapi_cursor` yields a cursor and drops the connection on a transient error. `dbapi_query_as_rows` returns Spark `Row`s keyed by column name. `execute_dbapi_write` runs one parameterised write and commits it. `get_next_id` takes `NEXTVAL`. `log_row_via_dbapi` inserts a log row and handles a primary-key collision. `is_own_committed_insert` tells "our commit already landed" from a real collision. `check_sequence_drift` is the startup warning. `read_snowflake`, `read_jdbc`, and `read_s3` are the three bulk readers. `check_full_load_row_count` is the extra S3 full-load check.

**Audit.** `run_guarded` and `never_raises` keep an audit or email failure from replacing the pipeline error. `audit_row_base` is the columns both log tables share. `insert_audit_row` allocates the id and inserts. `update_audit_row` updates by primary key. `log_job_event` and `log_object_event` insert. `update_job_log_status` and `update_object_log_status` move a row to `SUCCESS` or `FAILED`. `update_job_log_email_status` fills the `FAILURE_EMAIL_*` columns. `record_count_values` counts every written row as inserted. This job never updates individual target rows.

**Metadata.** `get_active_jobs`, `get_active_objects_bulk`, `get_last_watermarks_bulk`, and `get_all_failure_recipients`.

**Config.** `get_config_value` returns the stored value, or `None` when the column is absent. `get_config_int` reads the two partition knobs. Null means "leave the engine alone". `parse_source_schema` builds a Spark schema. `parse_target_s3_file_path` checks the four-part relative path.

**Processing.** `build_target_s3_path`, `build_target_file_name`, `derive_partition_column`, `build_source_query`, `get_latest_watermark`, `compute_target_partitions`, `write_to_s3`, `write_to_s3_with_named_file`, `delete_s3_prefix`, `check_object_preflight`, `resolve_jdbc_connection_name`, `extract_source`, `process_object`.

**Orchestration.** `html_escape`, `html_table`, and `build_failure_notification_message` build the mail. `send_email_notification` calls SES. `notify_run_failure` and `notify_job_failure` are the two senders. `build_object_result` is the dict the job rollup and the email share. `run_object`, `run_job_objects`, `run_pipeline`, `run_jobs_concurrently`, `run_all_jobs`.

**Entry.** `resolve_job_arguments`, `start_run`, `main`.

## Places the header and the code disagree

Worth knowing while you read, so a successful log line is not taken for a stricter check than the code performs.

- `DATA_LOAD_TYPE` is not checked against `{FULL, INCREMENTAL}`. Any value other than the exact string `INCREMENTAL`, including `full` and NULL, takes the unfiltered extract and can finish `SUCCESS`.
- `WRITE_MODE` is rejected when it is null. Spark accepts `OVERWRITE` because it lower-cases the mode, but the retry comparison is `== "overwrite"`, so that run does not retry a transient write error.
- The parquet check runs only when the extract is non-empty. An empty extract with another `TARGET_OBJECT_FILE_FORMAT` still returns success.
- The S3 full-load line count treats a quoted newline as extra rows, so a valid multiline CSV can fail the check. The check does not run when the parsed count is 0.
- Part files are copied to the final prefix one at a time. If a later part fails, earlier parts remain under the real date folder. Cleanup removes the `_tmp_` folder only.
- A non-empty incremental extract whose watermark column is all null stores a null watermark on `SUCCESS`. The next run does not see a resume point and extracts the full source again.
- `copy_object` cannot copy a single object over 5 GiB. A larger Spark part fails after the temp write.
- `read_s3` wraps every failure, including access denied and a null `SOURCE_S3_PATH`, with a lead sentence about an empty path. The original error is on the end of the same message.
