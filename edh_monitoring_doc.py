"""Generate EDH Operational Monitoring & Observability Word document."""

from docx import Document
from docx.shared import Pt, RGBColor, Cm, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy

doc = Document()

# ─── Page margins ─────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Cm(2.0)
    section.bottom_margin = Cm(2.0)
    section.left_margin   = Cm(2.5)
    section.right_margin  = Cm(2.5)

# ─── Style helpers ────────────────────────────────────────────────────────────

def set_cell_bg(cell, hex_color: str):
    """Set table cell background colour (hex without #)."""
    tc   = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd  = OxmlElement("w:shd")
    shd.set(qn("w:val"),   "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"),  hex_color)
    tcPr.append(shd)

def set_cell_borders(table):
    """Light grey borders on every cell."""
    for row in table.rows:
        for cell in row.cells:
            tc   = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                border = OxmlElement(f"w:{side}")
                border.set(qn("w:val"),   "single")
                border.set(qn("w:sz"),    "4")
                border.set(qn("w:space"), "0")
                border.set(qn("w:color"), "CCCCCC")
                tcBorders.append(border)
            tcPr.append(tcBorders)

def header_row(table, headers, bg="1F3864", fg="FFFFFF"):
    """Write a styled header row."""
    row = table.rows[0]
    for i, h in enumerate(headers):
        cell = row.cells[i]
        cell.text = h
        p    = cell.paragraphs[0]
        run  = p.runs[0]
        run.bold      = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor.from_string(fg)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        set_cell_bg(cell, bg)

def data_row(table, row_idx, values, shade=False, tone=None):
    """Write a data row with optional shading."""
    row = table.rows[row_idx]
    tone_colors = {
        "critical": "FDECEA",
        "high":     "FFF3E0",
        "warning":  "FFFDE7",
        "medium":   "E3F2FD",
        "success":  "E8F5E9",
        "info":     "E3F2FD",
    }
    bg = tone_colors.get(tone, "F5F5F5" if shade else "FFFFFF")
    for i, val in enumerate(values):
        cell = row.cells[i]
        cell.text = str(val)
        p    = cell.paragraphs[0]
        if p.runs:
            p.runs[0].font.size = Pt(9)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_cell_bg(cell, bg)

def add_table(headers, rows, col_widths=None, row_tones=None):
    """Create a formatted table and return it."""
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in t.rows:
                row.cells[i].width = Cm(w)
    header_row(t, headers)
    for ri, row_data in enumerate(rows):
        tone  = (row_tones[ri] if row_tones and ri < len(row_tones) else None)
        shade = ri % 2 == 1
        data_row(t, ri + 1, row_data, shade=shade, tone=tone)
    set_cell_borders(t)
    return t

def h1(text):
    p = doc.add_heading(text, level=1)
    p.runs[0].font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

def h2(text):
    p = doc.add_heading(text, level=2)
    p.runs[0].font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

def h3(text):
    p = doc.add_heading(text, level=3)
    p.runs[0].font.color.rgb = RGBColor(0x40, 0x40, 0x40)

def body(text, indent=False):
    p = doc.add_paragraph(text)
    p.runs[0].font.size = Pt(10) if p.runs else None
    if indent:
        p.paragraph_format.left_indent = Cm(0.5)
    return p

def bullet(text, level=0):
    p = doc.add_paragraph(text, style="List Bullet")
    p.runs[0].font.size = Pt(10) if p.runs else None
    p.paragraph_format.left_indent = Cm(0.5 + level * 0.5)
    return p

def note_box(text, label="Note"):
    """Callout-style paragraph."""
    p = doc.add_paragraph()
    run = p.add_run(f"  {label}:  ")
    run.bold = True
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    r2 = p.add_run(text)
    r2.font.size = Pt(9)
    p.paragraph_format.left_indent  = Cm(0.3)
    p.paragraph_format.right_indent = Cm(0.3)
    return p

def code_block(text):
    p = doc.add_paragraph(text)
    p.runs[0].font.name = "Courier New"
    p.runs[0].font.size = Pt(8)
    p.paragraph_format.left_indent = Cm(0.5)
    return p

def page_break():
    doc.add_page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# TITLE PAGE
# ═══════════════════════════════════════════════════════════════════════════════

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = p.add_run("\n\n\nEDH Operational Monitoring & Observability\nusing AWS CloudWatch")
run.bold = True
run.font.size = Pt(22)
run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

p2 = doc.add_paragraph()
p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
r2 = p2.add_run("Proof of Concept (PoC) — Implementation Guide\n")
r2.font.size = Pt(14)
r2.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

p3 = doc.add_paragraph()
p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
r3 = p3.add_run("Enterprise Data Hub (EDH) | Parent Story: US-1038931\n"
                  "Child Tasks: 1038932 · 1038933 · 1038935\n\n")
r3.font.size = Pt(11)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# TABLE OF CONTENTS (manual)
# ═══════════════════════════════════════════════════════════════════════════════

h1("Table of Contents")
toc_items = [
    ("1.", "Executive Summary"),
    ("2.", "Project Background & Scope"),
    ("3.", "EDH Architecture & Monitoring Layers"),
    ("4.", "Task 1038932 — Metrics & Monitoring"),
    ("  4.1", "Native AWS Metrics"),
    ("  4.2", "Custom EDH Metrics"),
    ("  4.3", "Recommended Dashboard Widgets"),
    ("5.", "Task 1038933 — CloudWatch Logs & Alarms"),
    ("  5.1", "Log Groups"),
    ("  5.2", "Logs Insights Saved Queries"),
    ("  5.3", "CloudWatch Alarms Configuration"),
    ("  5.4", "Alert Flow: Alarm → SNS → ServiceNow"),
    ("6.", "Custom EDH Metrics — Implementation Approaches"),
    ("  6.1", "Approach Comparison"),
    ("  6.2", "Approach 1 — Glue Script Code"),
    ("7.", "AWS Console Build Guide — Step by Step"),
    ("  7.1", "Phase 1: Enable Native Metrics"),
    ("  7.2", "Phase 2: IAM Permissions & Custom Metrics"),
    ("  7.3", "Phase 3: SNS Topic, Alarms & Dashboard"),
    ("  7.4", "Phase 4: Testing & PoC Validation"),
    ("8.", "Test Scenarios"),
    ("9.", "Task 1038935 — PoC Evaluation"),
    ("  9.1", "Evaluation Scores"),
    ("  9.2", "Technical Limitations"),
    ("  9.3", "Recommendations for EDH-Wide Adoption"),
    ("10.", "IAM Permissions Reference"),
]
for num, title in toc_items:
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Cm(0.5) if num.startswith(" ") else Cm(0)
    run = p.add_run(f"{num.strip():<6}{title}")
    run.font.size = Pt(10)
    if not num.startswith(" "):
        run.bold = True

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 1. EXECUTIVE SUMMARY
# ═══════════════════════════════════════════════════════════════════════════════

h1("1. Executive Summary")
body(
    "This document is the implementation guide for the EDH Operational Monitoring and Observability "
    "Proof of Concept (PoC) using AWS CloudWatch. It covers the investigation, design, console-based "
    "implementation, testing, and evaluation of centralised operational monitoring for the Enterprise "
    "Data Hub (EDH) architecture."
)
body(
    "The PoC evaluates AWS CloudWatch as the operational monitoring platform for EDH pipelines, "
    "specifically for AWS Glue ingestion jobs, Step Functions orchestration, S3 data storage, "
    "EventBridge scheduling, and Snowflake downstream processing."
)
doc.add_paragraph()
note_box(
    "All implementation activities described in this document are performed exclusively through the "
    "AWS Management Console. No Terraform, CDK, or infrastructure-as-code is used.",
    label="Important"
)
doc.add_paragraph()

add_table(
    ["Area", "Finding", "Recommendation"],
    [
        ["Overall CloudWatch Suitability", "3.7 / 5 — Suitable for EDH operational monitoring", "Proceed with full implementation in DEV, then TEST, then PROD"],
        ["Native AWS Metrics", "19 metrics available across Glue, Step Functions, S3, EventBridge", "Enable immediately — no code changes required"],
        ["Custom EDH Metrics", "10 custom metrics required across 4 namespaces", "Implement via PutMetricData in Glue job script"],
        ["Alerting", "CloudWatch Alarms + SNS covers all critical failure paths", "Create 10 alarms; ServiceNow integration via HTTPS subscription"],
        ["Implementation Effort", "Phase 1: Low. Phase 2–3: Medium. ServiceNow: High", "Prioritise Phases 1–3; ServiceNow in a separate sprint"],
    ],
    col_widths=[4.5, 7.5, 6.5],
    row_tones=["info", "success", "info", "success", "warning"]
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 2. PROJECT BACKGROUND & SCOPE
# ═══════════════════════════════════════════════════════════════════════════════

h1("2. Project Background & Scope")
h2("2.1 Existing EDH Architecture")
body("The EDH architecture uses the following components:")
for item in [
    "AWS Glue — data ingestion and validation (job: edh-glue-ingestion-daily)",
    "Amazon S3 — storing ingested data in Parquet format (EDH data lake)",
    "Snowflake — downstream data processing and orchestration",
    "AWS Step Functions — workflow orchestration across the pipeline",
    "Amazon EventBridge — scheduling and triggering workflows",
    "Amazon SES — failure notifications",
    "Snowflake metadata and validation tables — job execution logs, object logs, validation logs, task DAG logs",
]:
    bullet(item)

h2("2.2 Existing Ingestion Workflow")
body("The pipeline follows this sequence:")
bullet("Source System → AWS Glue Ingestion (edh-glue-ingestion-daily)")
bullet("Glue → S3 (Parquet files written to data lake)")
bullet("Glue → Validation (executed within the Glue ingestion process)")
bullet("Validation Success → Snowflake Root Task → Downstream Task Execution")
bullet("Validation Failure / Glue Failure → SES Notification via Step Functions")

h2("2.3 PoC Scope — Child Tasks")
add_table(
    ["Task ID", "Title", "Scope"],
    [
        ["1038932", "Metrics & Monitoring", "Native CloudWatch metrics for all AWS services. Custom EDH metric design and publishing. Dashboard design."],
        ["1038933", "Logs & Alerts", "CloudWatch Log groups, Logs Insights queries, CloudWatch alarms, SNS notifications, ServiceNow integration assessment."],
        ["1038935", "Evaluation", "Scored assessment of CloudWatch suitability, effort, limitations, and enterprise-wide adoption recommendations."],
    ],
    col_widths=[2.5, 4.5, 11.5],
    row_tones=["info", "info", "info"]
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 3. ARCHITECTURE & MONITORING LAYERS
# ═══════════════════════════════════════════════════════════════════════════════

h1("3. EDH Architecture & Monitoring Layers")
body(
    "The monitoring architecture overlays three complementary layers on top of the existing EDH pipeline. "
    "Each layer addresses a different aspect of observability."
)

h2("3.1 Monitoring Layer Summary")
add_table(
    ["Monitoring Layer", "Scope", "AWS Service", "Availability"],
    [
        ["Infrastructure Monitoring", "Glue, S3, Step Functions, EventBridge", "CloudWatch Metrics", "Native — no code change required"],
        ["Application Monitoring", "Record counts, validation results, pipeline status", "CloudWatch Custom Metrics", "Requires PutMetricData calls in Glue script"],
        ["Log-based Observability", "Job execution logs, errors, structured events", "CloudWatch Logs + Logs Insights", "Partial auto (Glue); custom log group requires code"],
        ["Alerting & Notification", "Critical failures, threshold breaches", "CloudWatch Alarms + SNS", "Requires alarm + SNS configuration"],
        ["Incident Management", "ServiceNow ticket creation from operational alerts", "SNS HTTPS → ServiceNow", "Requires ServiceNow webhook setup"],
    ],
    col_widths=[4, 5, 4, 5.5],
    row_tones=["neutral", "info", "neutral", "neutral", "warning"]
)

h2("3.2 Pipeline Flow and Monitoring Points")
body("The EDH pipeline and its associated monitoring coverage:")
monitoring_points = [
    ("EventBridge", "Scheduled Trigger", "CloudWatch Metrics: Invocations, FailedInvocations, ThrottledRules"),
    ("Step Functions", "Workflow Orchestrator", "CloudWatch Metrics: ExecutionsFailed, ExecutionTime. CloudWatch Logs: ALL state transitions"),
    ("AWS Glue (edh-glue-ingestion-daily)", "Data Ingestion + Validation", "CloudWatch Metrics (native): 7 metrics. CloudWatch Custom Metrics: 9 EDH metrics via PutMetricData. CloudWatch Logs: output + error + custom structured logs"),
    ("Amazon S3", "Parquet Data Lake", "CloudWatch Metrics: BucketSizeBytes, NumberOfObjects, 5xxErrors, FirstByteLatency"),
    ("Snowflake Task", "Root Task + DAG", "Custom Metrics only: DownstreamTaskStatus published via intermediate Lambda or Glue step"),
    ("Downstream Processing", "Snowflake Processing", "Custom Metrics: published from EDH Snowflake metadata tables"),
    ("Amazon SES", "Failure Notifications", "No CloudWatch metrics for SES in this PoC scope"),
]
for svc, role, monitoring in monitoring_points:
    p = doc.add_paragraph()
    r1 = p.add_run(f"{svc}  ")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p.add_run(f"({role})\n")
    r2.font.size = Pt(10)
    r2.italic = True
    r3 = p.add_run(f"  Monitoring: {monitoring}")
    r3.font.size = Pt(9)
    p.paragraph_format.left_indent = Cm(0.5)
    p.paragraph_format.space_after = Pt(4)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 4. TASK 1038932 — METRICS & MONITORING
# ═══════════════════════════════════════════════════════════════════════════════

h1("4. Task 1038932 — Metrics & Monitoring")
note_box(
    "Native metrics are automatically published by AWS to CloudWatch — you only need to enable them "
    "in the service console. Custom EDH metrics must be explicitly published from the Glue job using "
    "boto3 PutMetricData calls.",
    label="Key Distinction"
)
doc.add_paragraph()

h2("4.1 Native AWS Metrics")
body("The following 19 native CloudWatch metrics are relevant to the EDH pipeline. No code changes are required.")

h3("AWS Glue — 7 Metrics")
body("Console navigation to enable: Glue → ETL Jobs → edh-glue-ingestion-daily → Job details → Advanced properties → Toggle ON 'Job metrics' and 'Continuous logging'.")
add_table(
    ["Metric Name", "Description", "Setup Required"],
    [
        ["glue.driver.aggregate.numCompletedTasks", "Total Spark tasks completed successfully", "Enable 'Job metrics' in job settings"],
        ["glue.driver.aggregate.numFailedTasks",    "Total Spark tasks that failed",             "Enable 'Job metrics' in job settings"],
        ["glue.driver.aggregate.numKilledTasks",    "Total Spark tasks that were killed",        "Enable 'Job metrics' in job settings"],
        ["glue.driver.aggregate.recordsRead",       "Total records read across all DPUs",        "Enable 'Job metrics' in job settings"],
        ["glue.driver.aggregate.bytesRead",         "Total bytes read from all data sources",    "Enable 'Job metrics' in job settings"],
        ["glue.ALL.jvm.heap.usage",                 "JVM heap memory usage ratio (0–1.0)",       "Enable 'Job metrics' in job settings"],
        ["glue.driver.ExecutorRunTime",             "Total executor execution time (ms)",        "Enable 'Job metrics' in job settings"],
    ],
    col_widths=[7, 6, 5.5],
    row_tones=["success"] * 7
)

doc.add_paragraph()
h3("AWS Step Functions — 5 Metrics")
body("No configuration required. Metrics are published automatically for all State Machines.")
add_table(
    ["Metric Name", "Description", "Setup Required"],
    [
        ["ExecutionsStarted",   "Executions initiated during the period",           "No configuration required"],
        ["ExecutionsFailed",    "Executions that ended in FAILED state",             "No configuration required"],
        ["ExecutionsSucceeded", "Executions that completed successfully",            "No configuration required"],
        ["ExecutionsTimedOut",  "Executions that exceeded the configured timeout",   "No configuration required"],
        ["ExecutionTime",       "Duration from start to completion in milliseconds", "No configuration required"],
    ],
    col_widths=[5, 7, 6.5],
    row_tones=["info"] * 5
)

doc.add_paragraph()
h3("Amazon S3 — 4 Metrics")
body("Storage metrics (BucketSizeBytes, NumberOfObjects) are published daily at no extra setup. Request metrics require enabling a filter.")
add_table(
    ["Metric Name", "Description", "Setup Required"],
    [
        ["BucketSizeBytes",    "Total size of objects in the bucket",    "Enable Storage Lens or storage class metrics"],
        ["NumberOfObjects",    "Total count of objects stored",          "Enable Storage Lens or storage class metrics"],
        ["5xxErrors",          "S3 server-side error rate on requests",  "Enable Request Metrics filter on bucket"],
        ["FirstByteLatency",   "Per-request latency for first byte",     "Enable Request Metrics filter on bucket"],
    ],
    col_widths=[5, 7, 6.5],
)

doc.add_paragraph()
h3("Amazon EventBridge — 3 Metrics")
body("No configuration required. Metrics are published automatically for all EventBridge rules.")
add_table(
    ["Metric Name", "Description", "Setup Required"],
    [
        ["Invocations",       "Rule target invocations triggered",           "No configuration required"],
        ["FailedInvocations", "Invocations that could not be delivered",     "No configuration required"],
        ["ThrottledRules",    "Rules throttled during invocation",           "No configuration required"],
    ],
    col_widths=[5, 7, 6.5],
)

page_break()

h2("4.2 Custom EDH Metrics (Require PutMetricData)")
body(
    "The following 10 custom metrics do not exist in CloudWatch by default. They must be published "
    "from the Glue job using boto3 PutMetricData calls. Use the EDH/ namespace prefix to separate "
    "these from AWS native metrics."
)
add_table(
    ["Namespace", "Metric Name", "Description", "Unit", "Dimensions"],
    [
        ["EDH/Ingestion",  "RecordsProcessed",        "Total records successfully written to S3",         "Count",   "JobName, PipelineName, Environment"],
        ["EDH/Ingestion",  "RecordsRejected",          "Records rejected during validation",               "Count",   "JobName, PipelineName, Environment"],
        ["EDH/Ingestion",  "SourceRecordCount",        "Total records read from source system",            "Count",   "JobName, SourceSystem, Environment"],
        ["EDH/Ingestion",  "TargetRecordCount",        "Total records written to target",                  "Count",   "JobName, TargetSystem, Environment"],
        ["EDH/Validation", "ValidationPassCount",      "Validation checks that passed",                   "Count",   "JobName, ObjectName, Environment"],
        ["EDH/Validation", "ValidationFailCount",      "Validation checks that failed",                   "Count",   "JobName, ObjectName, Environment"],
        ["EDH/Pipeline",   "PipelineExecutionStatus",  "1 = Success, 0 = Failure",                        "None",    "PipelineName, Environment"],
        ["EDH/Pipeline",   "JobExecutionDuration",     "Total job runtime in seconds",                    "Seconds", "JobName, Environment"],
        ["EDH/Pipeline",   "RecordCountMismatch",      "Abs. difference between source and target counts","Count",   "JobName, ObjectName, Environment"],
        ["EDH/Snowflake",  "DownstreamTaskStatus",     "Snowflake downstream task status (1=OK, 0=Fail)", "None",    "TaskName, Environment"],
    ],
    col_widths=[3.5, 4.5, 5.5, 2, 3],
)

doc.add_paragraph()
h2("4.3 Recommended CloudWatch Dashboard Widgets")
add_table(
    ["Widget Type", "Metric / Source", "Purpose"],
    [
        ["Number (single value)", "ExecutionsFailed (Step Functions)",               "Failed executions count — instant pipeline health check"],
        ["Number (single value)", "PipelineExecutionStatus (EDH/Pipeline)",          "1=Success / 0=Failure — current pipeline state"],
        ["Line chart",            "ExecutionTime (Step Functions)",                  "Execution duration trend over time"],
        ["Line chart",            "RecordsProcessed vs RecordsRejected",             "Ingestion quality trend"],
        ["Stacked area",          "ValidationPassCount vs ValidationFailCount",      "Validation health over time"],
        ["Line chart",            "SourceRecordCount vs TargetRecordCount",          "Record count reconciliation"],
        ["Alarm status widget",   "All EDH-* alarms",                               "Single-pane alarm health overview"],
        ["Number",                "glue.ALL.jvm.heap.usage",                        "Glue memory pressure indicator"],
    ],
    col_widths=[4, 6, 8.5],
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 5. TASK 1038933 — LOGS & ALARMS
# ═══════════════════════════════════════════════════════════════════════════════

h1("5. Task 1038933 — CloudWatch Logs & Alarms")

h2("5.1 CloudWatch Log Groups")
note_box(
    "Console navigation: CloudWatch → Log groups. Set retention: select log group → Actions → "
    "Edit retention → 30 days. Create new group: CloudWatch → Log groups → Create log group.",
    label="Console Navigation"
)
doc.add_paragraph()
add_table(
    ["Log Group Name", "Service", "Description", "How Created", "Retention"],
    [
        ["/aws-glue/jobs/output",           "Glue",         "Standard output from Glue job drivers and executors",          "Auto (enable Continuous logging in Glue job settings)",       "Set to 30 days"],
        ["/aws-glue/jobs/error",            "Glue",         "Exception traces and error messages from Glue jobs",           "Auto (enable Continuous logging in Glue job settings)",       "Set to 30 days"],
        ["/aws-glue/jobs/logs-v2",          "Glue",         "Continuous logging output (Glue 3.0+)",                       "Auto (enable Continuous logging in Glue job settings)",       "Set to 30 days"],
        ["/aws/states/edh-step-functions",  "Step Fns",     "Execution logs: state transitions, input/output, errors",     "Manual: configure on State Machine logging → Level: ALL",    "Set to 30 days"],
        ["/edh/custom/ingestion",           "Glue (custom)","Structured JSON application logs from EDH Glue scripts",      "Manual: create log group + configure in Glue script",        "Set to 30 days"],
    ],
    col_widths=[5.5, 2, 5, 4.5, 1.5],
)

doc.add_paragraph()
h2("5.2 Logs Insights Saved Queries")
note_box(
    "Console navigation: CloudWatch → Logs Insights → Select log group → Paste query → Run query → "
    "Actions → Save query. Saved queries appear under Queries in the left panel.",
    label="Console Navigation"
)
doc.add_paragraph()

queries = [
    {
        "title": "Query 1 — Glue Job Failures (last 24 hours)",
        "log_group": "/aws-glue/jobs/error",
        "purpose": "Find all Glue job errors, exceptions, and failures in the last 24 hours.",
        "code": (
            "fields @timestamp, @message\n"
            "| filter @message like /ERROR|Exception|FAILED/\n"
            "| sort @timestamp desc\n"
            "| limit 50"
        ),
    },
    {
        "title": "Query 2 — EDH Record Counts from Custom Logs",
        "log_group": "/edh/custom/ingestion",
        "purpose": "Extract source, target, and rejected record counts from EDH structured logs.",
        "code": (
            "fields @timestamp, jobName, sourceCount, targetCount, rejectedCount\n"
            "| filter ispresent(sourceCount)\n"
            "| sort @timestamp desc\n"
            "| limit 100"
        ),
    },
    {
        "title": "Query 3 — Step Functions Execution Failures",
        "log_group": "/aws/states/edh-step-functions",
        "purpose": "Find failed executions and task failures with error details and cause.",
        "code": (
            "fields @timestamp, execution_arn, type, details.error, details.cause\n"
            "| filter type = \"ExecutionFailed\" or type = \"TaskFailed\"\n"
            "| sort @timestamp desc\n"
            "| limit 50"
        ),
    },
    {
        "title": "Query 4 — Validation Failures by Object Name",
        "log_group": "/edh/custom/ingestion",
        "purpose": "Group validation failures by object name to identify the most problematic objects.",
        "code": (
            "fields @timestamp, jobName, objectName, validationRule, failureReason\n"
            "| filter validationStatus = \"FAIL\"\n"
            "| stats count(*) as failureCount by objectName\n"
            "| sort failureCount desc"
        ),
    },
    {
        "title": "Query 5 — Average Glue Job Duration (Daily Trend)",
        "log_group": "/aws-glue/jobs/output",
        "purpose": "Calculate average and maximum job execution duration per day for trend analysis.",
        "code": (
            "fields @timestamp, @message\n"
            "| filter @message like /Job run time/\n"
            "| parse @message \"Job run time: * seconds\" as durationSec\n"
            "| stats avg(durationSec) as avgDuration, max(durationSec) as maxDuration by bin(1d)"
        ),
    },
]

for q in queries:
    h3(q["title"])
    body(f"Log Group: {q['log_group']}")
    body(f"Purpose: {q['purpose']}")
    code_block(q["code"])
    doc.add_paragraph()

page_break()

h2("5.3 CloudWatch Alarms Configuration")
note_box(
    "Console navigation: CloudWatch → Alarms → All alarms → Create alarm → Select metric → "
    "Configure threshold and period → Set alarm state action → Select SNS topic "
    "(edh-operational-alerts) → Add alarm name → Create alarm.",
    label="Console Navigation"
)
doc.add_paragraph()

add_table(
    ["Alarm Name", "Metric", "Namespace", "Threshold", "Period", "SNS Action", "Priority"],
    [
        ["EDH-GlueJob-TaskFailure",          "glue.driver.aggregate.numFailedTasks",  "Glue",   "> 0",           "5 min",  "edh-operational-alerts", "Critical"],
        ["EDH-StepFunctions-ExecutionFailed","ExecutionsFailed",                       "States", ">= 1",          "5 min",  "edh-operational-alerts", "Critical"],
        ["EDH-Custom-PipelineFailure",       "PipelineExecutionStatus (EDH/Pipeline)", "EDH",    "< 1 (= 0)",     "5 min",  "edh-operational-alerts", "Critical"],
        ["EDH-StepFunctions-ExecutionTimeout","ExecutionsTimedOut",                    "States", ">= 1",          "5 min",  "edh-operational-alerts", "High"],
        ["EDH-StepFunctions-LongRunning",    "ExecutionTime",                          "States", "> 3,600,000 ms","15 min", "edh-operational-alerts", "High"],
        ["EDH-Custom-ValidationFailure",     "ValidationFailCount (EDH/Validation)",   "EDH",    "> 0",           "5 min",  "edh-operational-alerts", "High"],
        ["EDH-EventBridge-FailedInvocation", "FailedInvocations",                      "Events", ">= 1",          "5 min",  "edh-operational-alerts", "High"],
        ["EDH-Custom-RecordMismatch",        "RecordCountMismatch (EDH/Pipeline)",      "EDH",    "> 100",         "5 min",  "edh-operational-alerts", "Medium"],
        ["EDH-S3-ServerErrors",              "5xxErrors",                               "S3",     "> 5",           "5 min",  "edh-operational-alerts", "Medium"],
        ["EDH-Glue-HighMemory",              "glue.ALL.jvm.heap.usage",                 "Glue",   "> 0.85 (85%)",  "5 min",  "edh-operational-alerts", "Warning"],
    ],
    col_widths=[5, 5, 1.8, 2.5, 1.5, 2.5, 1.7],
    row_tones=["critical","critical","critical","high","high","high","high","medium","medium","warning"]
)

doc.add_paragraph()
h2("5.4 Alert Flow: Alarm → SNS → ServiceNow")
body(
    "The following five-stage flow describes how a metric threshold breach progresses from detection "
    "to incident creation."
)
add_table(
    ["Stage", "Component", "Role", "Console Location"],
    [
        ["1. Detection",         "CloudWatch Alarm",       "Evaluates metric against threshold; transitions to ALARM state",                      "CloudWatch → Alarms → All alarms"],
        ["2. Notification",      "Amazon SNS",             "Distributes alert to all subscribers (email, HTTPS endpoints)",                       "SNS → Topics → edh-operational-alerts"],
        ["3. Email Alert",       "SNS Email Subscription", "Sends immediate email to operations team",                                            "SNS → Subscriptions → Protocol: Email"],
        ["4. Incident Creation", "ServiceNow Webhook",     "HTTPS POST to ServiceNow inbound integration creates an incident ticket",             "SNS → Subscriptions → Protocol: HTTPS → ServiceNow URL"],
        ["5. Resolution",        "CloudWatch Alarm",       "Transitions to OK state when metric recovers; sends resolution notification to SNS",  "CloudWatch → Alarms → alarm → History tab"],
    ],
    col_widths=[2.5, 3.5, 7, 5.5],
)
doc.add_paragraph()
note_box(
    "Monitoring, Alerting, and Incident Creation are three separate stages. "
    "Monitoring = CloudWatch collects metrics. "
    "Alerting = CloudWatch Alarm fires and notifies SNS. "
    "Incident Creation = ServiceNow webhook creates a ticket. Each stage can be configured and tested independently.",
    label="Key Distinction"
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 6. CUSTOM EDH METRICS — APPROACHES
# ═══════════════════════════════════════════════════════════════════════════════

h1("6. Custom EDH Metrics — Implementation Approaches")
body(
    "The EDH framework already maintains metadata and validation results in Snowflake tables. "
    "Three approaches can surface those values as CloudWatch custom metrics."
)

h2("6.1 Approach Comparison")
add_table(
    ["Approach", "Latency", "New Infrastructure", "Effort", "Best For"],
    [
        ["1. PutMetricData directly from Glue job (Recommended)", "Real-time",           "None",                   "Low–Medium",   "Glue-level metrics: records, validation, duration, pipeline status"],
        ["2. Intermediate Lambda reading Snowflake metadata",      "Minutes (async)",     "Lambda + Secrets Mgr",  "Medium–High",  "Snowflake downstream task status and validation results"],
        ["3. Scheduled metric publisher (EventBridge → Lambda)",   "Schedule interval",   "EventBridge + Lambda",  "Medium",       "Reporting dashboards only — not for real-time alarms"],
    ],
    col_widths=[5.5, 2, 3, 2.5, 5.5],
    row_tones=["success", "info", "neutral"]
)

doc.add_paragraph()
h3("Approach 1 Advantages")
for pro in [
    "Lowest latency — metrics published as the pipeline runs in real time",
    "No intermediate infrastructure required",
    "Uses existing Glue IAM role (only add cloudwatch:PutMetricData permission)",
    "Single code change inside the Glue job script",
]:
    bullet(f"+ {pro}")

h3("Approach 1 Disadvantages")
for con in [
    "Glue script changes are required — involves a code change and deployment",
    "If Glue crashes early in execution, some metrics may not be published",
    "Metrics are only available when the Glue job runs",
]:
    bullet(f"- {con}")

doc.add_paragraph()
h2("6.2 Approach 1 — Glue Script Code (PutMetricData)")
note_box(
    "Before adding this code: IAM → Roles → (Glue job IAM role) → Add permissions → "
    "Create inline policy → JSON tab → add cloudwatch:PutMetricData → save as "
    "EDH-CloudWatch-Custom-Metrics. See Section 7.2, Step 2.1 for the full IAM JSON.",
    label="IAM Prerequisite"
)
doc.add_paragraph()
body("Add the following function to edh-glue-ingestion-daily Glue job script. Call it after validation is complete.")
code_block("""import boto3

def publish_edh_metrics(job_name, pipeline_name, environment,
                        records_processed, records_rejected,
                        source_count, target_count,
                        validation_pass, validation_fail,
                        duration_sec):
    cw = boto3.client('cloudwatch', region_name='ap-southeast-2')
    base_dims = [
        {'Name': 'JobName',      'Value': job_name},
        {'Name': 'PipelineName', 'Value': pipeline_name},
        {'Name': 'Environment',  'Value': environment},
    ]

    # EDH/Ingestion namespace
    cw.put_metric_data(Namespace='EDH/Ingestion', MetricData=[
        {'MetricName': 'RecordsProcessed',  'Value': records_processed, 'Unit': 'Count',   'Dimensions': base_dims},
        {'MetricName': 'RecordsRejected',   'Value': records_rejected,  'Unit': 'Count',   'Dimensions': base_dims},
        {'MetricName': 'SourceRecordCount', 'Value': source_count,      'Unit': 'Count',   'Dimensions': base_dims},
        {'MetricName': 'TargetRecordCount', 'Value': target_count,      'Unit': 'Count',   'Dimensions': base_dims},
    ])

    # EDH/Validation namespace
    cw.put_metric_data(Namespace='EDH/Validation', MetricData=[
        {'MetricName': 'ValidationPassCount', 'Value': validation_pass, 'Unit': 'Count', 'Dimensions': base_dims},
        {'MetricName': 'ValidationFailCount', 'Value': validation_fail, 'Unit': 'Count', 'Dimensions': base_dims},
    ])

    # EDH/Pipeline namespace
    mismatch = abs(source_count - target_count)
    status   = 1 if validation_fail == 0 and mismatch == 0 else 0
    cw.put_metric_data(Namespace='EDH/Pipeline', MetricData=[
        {'MetricName': 'PipelineExecutionStatus', 'Value': status,       'Unit': 'None',    'Dimensions': base_dims},
        {'MetricName': 'JobExecutionDuration',    'Value': duration_sec, 'Unit': 'Seconds', 'Dimensions': base_dims},
        {'MetricName': 'RecordCountMismatch',     'Value': mismatch,     'Unit': 'Count',   'Dimensions': base_dims},
    ])


# --- Call after validation is complete (replace with your actual variable names) ---
publish_edh_metrics(
    job_name='edh-glue-ingestion-daily',
    pipeline_name='edh-daily',
    environment='dev',                   # change to 'test' or 'prod' per environment
    records_processed=total_records,
    records_rejected=rejected_records,
    source_count=source_record_count,
    target_count=target_record_count,
    validation_pass=validation_pass_count,
    validation_fail=validation_fail_count,
    duration_sec=job_duration_seconds
)""")

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 7. BUILD GUIDE — CONSOLE STEP BY STEP
# ═══════════════════════════════════════════════════════════════════════════════

h1("7. AWS Console Build Guide — Step by Step")
body(
    "This section provides the exact AWS Management Console navigation, field values, and "
    "validation steps for every implementation activity. Work through phases in order."
)
note_box(
    "All activities must be performed in the DEV environment first. Do not apply changes to TEST "
    "or PROD without completing validation in DEV.",
    label="Important"
)
doc.add_paragraph()

add_table(
    ["Phase", "Steps", "IAM Changes Required", "Estimated Time"],
    [
        ["Phase 1 — Enable Native Metrics",    "1.1 → 1.4", "None",                                       "~20 minutes"],
        ["Phase 2 — Custom Metrics",           "2.1 → 2.4", "Add cloudwatch:PutMetricData to Glue role",  "~45 minutes"],
        ["Phase 3 — Alarms & Dashboard",       "3.1 → 3.4", "None (console only)",                        "~60 minutes"],
        ["Phase 4 — Testing & Validation",     "4.1 → 4.2", "None",                                       "~60 minutes"],
    ],
    col_widths=[5, 2, 7, 4.5],
    row_tones=["success", "info", "warning", "neutral"]
)

doc.add_paragraph()

# ─── Phase 1 ──────────────────────────────────────────────────────────────────

h2("7.1 Phase 1 — Enable Native Metrics (No Code Changes)")
note_box("These steps require only AWS Console access. No code changes or IAM modifications needed.", label="Note")
doc.add_paragraph()

steps_p1 = [
    {
        "id": "1.1",
        "title": "Enable CloudWatch Metrics in the Glue Job",
        "objective": "Activates native CloudWatch metrics for edh-glue-ingestion-daily: task completion counts, records read, bytes read, and JVM heap usage.",
        "service": "AWS Glue",
        "time": "5 minutes",
        "steps": [
            "Open AWS Console → top search bar → type Glue → click AWS Glue",
            "In the left navigation panel, click ETL Jobs",
            "In the job list, click the name edh-glue-ingestion-daily",
            "Click the Job details tab (not the Script tab)",
            "Scroll down to the Advanced properties section",
            "Under Monitoring options, toggle ON Job metrics",
            "Toggle ON Continuous logging",
            "Continuous logging → Log group field: leave blank (auto-creates /aws-glue/jobs)",
            "Click Save job details button at the top right",
        ],
        "config": [
            ("Job metrics toggle", "ON"),
            ("Continuous logging toggle", "ON"),
            ("Log group (auto-created)", "/aws-glue/jobs/output and /aws-glue/jobs/error"),
            ("Applies from", "Next job run onwards"),
        ],
        "output": "Settings saved. After next job run, metrics appear: CloudWatch → Metrics → AWS Glue → Job Metrics.",
        "validate": "Run job once → CloudWatch → Metrics → All metrics → AWS Glue → Job Metrics → select edh-glue-ingestion-daily → verify data points on graph.",
        "tip": "If the Job metrics toggle is not visible, your Glue job may be using version 1.0. Update to Glue 2.0 or above: Job details → Glue version dropdown.",
    },
    {
        "id": "1.2",
        "title": "Enable Execution Logging on the EDH Step Functions State Machine",
        "objective": "Captures every state transition, error, and input/output value from the EDH workflow into a CloudWatch Log group.",
        "service": "AWS Step Functions",
        "time": "5 minutes",
        "steps": [
            "AWS Console → search bar → type Step Functions → click",
            "In left panel, click State machines",
            "Click the name of your EDH state machine",
            "Click the Edit button at the top right of the page",
            "Scroll down to the Logging section",
            "Logging level dropdown: select ALL (not ERROR — ALL captures all transitions)",
            "CloudWatch log group: click Create new log group button",
            "Log group name field: type /aws/states/edh-step-functions → click Create",
            "Toggle ON Include execution data",
            "X-Ray tracing: leave disabled for PoC scope",
            "Click Save button",
        ],
        "config": [
            ("Logging level", "ALL"),
            ("Log group name", "/aws/states/edh-step-functions"),
            ("Include execution data", "ON"),
            ("X-Ray tracing", "Disabled"),
        ],
        "output": "Log group created. After next execution, one log stream per workflow execution appears.",
        "validate": "Run the EDH pipeline → CloudWatch → Log groups → /aws/states/edh-step-functions → verify log stream exists.",
        "tip": "If Save fails with an IAM error, the state machine role needs: logs:CreateLogDelivery, logs:GetLogDelivery, logs:UpdateLogDelivery, logs:DescribeLogGroups, logs:PutResourcePolicy.",
    },
    {
        "id": "1.3",
        "title": "Enable Request Metrics on the EDH S3 Bucket",
        "objective": "Enables per-request S3 metrics such as 5xxErrors and FirstByteLatency.",
        "service": "Amazon S3",
        "time": "5 minutes",
        "steps": [
            "AWS Console → search bar → type S3 → click",
            "In the bucket list, click your EDH data lake bucket (e.g. edh-data-lake-dev)",
            "Click the Metrics tab",
            "Scroll to Request metrics section",
            "Click Create filter button",
            "Filter name field: type edh-all-requests",
            "Filter scope: select This filter applies to all objects in the bucket",
            "Click Create filter button",
            "Metrics will begin appearing in CloudWatch after approximately 15 minutes",
        ],
        "config": [
            ("Filter name", "edh-all-requests"),
            ("Scope", "Entire bucket (all objects)"),
            ("Metrics available after", "~15 minutes of first request activity"),
            ("CloudWatch namespace", "AWS/S3"),
        ],
        "output": "Filter created. After 15 minutes, CloudWatch → Metrics → AWS/S3 → edh-all-requests → 5xxErrors visible.",
        "validate": "CloudWatch → Metrics → AWS/S3 → by Request Metrics filter → edh-all-requests → verify 5xxErrors and AllRequests are present.",
        "tip": "S3 request metrics have a small additional CloudWatch cost (~$0.01 per 1,000 requests). Confirm with your AWS team.",
    },
    {
        "id": "1.4",
        "title": "Set 30-Day Retention on All EDH Log Groups",
        "objective": "By default CloudWatch log groups have Never expire retention — cost accumulates indefinitely. 30-day retention is the best practice.",
        "service": "Amazon CloudWatch",
        "time": "5 minutes",
        "steps": [
            "AWS Console → CloudWatch → Log groups (left panel)",
            "In the search box, type /aws-glue to filter",
            "Check the checkbox next to /aws-glue/jobs/output",
            "Click Actions dropdown → Edit retention setting",
            "Retention dropdown: select 30 days (1 month) → click Save changes",
            "Repeat for /aws-glue/jobs/error",
            "Repeat for /aws-glue/jobs/logs-v2 if present",
            "Clear the search box → type /aws/states to find the Step Functions log group",
            "Repeat for /aws/states/edh-step-functions → set to 30 days → Save",
        ],
        "config": [
            ("/aws-glue/jobs/output", "30 days"),
            ("/aws-glue/jobs/error", "30 days"),
            ("/aws-glue/jobs/logs-v2", "30 days"),
            ("/aws/states/edh-step-functions", "30 days"),
        ],
        "output": "Log groups list → Retention column shows '1 month (30 days)' for all /aws-glue/* groups.",
        "validate": "CloudWatch → Log groups → verify Retention column is not 'Never expire' for any EDH log group.",
        "tip": "If Actions → Edit retention is greyed out, you need logs:PutRetentionPolicy IAM permission.",
    },
]

steps_p2 = [
    {
        "id": "2.1",
        "title": "Add cloudwatch:PutMetricData Permission to the Glue IAM Role",
        "objective": "The Glue job runs under an IAM role that by default cannot write to CloudWatch. This grants the minimum permissions required.",
        "service": "AWS IAM",
        "time": "10 minutes",
        "steps": [
            "AWS Console → search bar → type IAM → click",
            "To find your Glue role name: Glue → ETL Jobs → edh-glue-ingestion-daily → Job details tab → scroll to IAM role field",
            "Return to IAM → left panel → Roles → paste the role name in the search box → click it",
            "Under Permissions tab → click Add permissions → Create inline policy",
            "Click the JSON tab at the top of the policy editor",
            "Delete all existing text in the editor",
            "Paste the IAM policy JSON (see Config Values below)",
            "Click Next",
            "Policy name field: type EDH-CloudWatch-Custom-Metrics",
            "Click Create policy",
        ],
        "config": [
            ("IAM Policy JSON", '{\n  "Version": "2012-10-17",\n  "Statement": [{\n    "Effect": "Allow",\n    "Action": [\n      "cloudwatch:PutMetricData",\n      "logs:CreateLogGroup",\n      "logs:CreateLogStream",\n      "logs:PutLogEvents"\n    ],\n    "Resource": "*"\n  }]\n}'),
            ("Policy name", "EDH-CloudWatch-Custom-Metrics"),
            ("Policy type", "Inline policy on the Glue job IAM role"),
        ],
        "output": "New inline policy EDH-CloudWatch-Custom-Metrics appears under the role's Permissions tab.",
        "validate": "IAM → Roles → your Glue role → Permissions tab → verify EDH-CloudWatch-Custom-Metrics is listed with cloudwatch:PutMetricData.",
        "tip": "Attach this policy only to the role used by edh-glue-ingestion-daily to follow least-privilege principles.",
    },
    {
        "id": "2.2",
        "title": "Update the Glue Job Script with Custom Metric Publishing",
        "objective": "Adds the publish_edh_metrics() function to the existing Glue job and calls it after validation completes.",
        "service": "AWS Glue",
        "time": "20–30 minutes",
        "steps": [
            "Glue → ETL Jobs → click edh-glue-ingestion-daily",
            "Click the Script tab",
            "Verify import boto3 is in the imports section at the top; add it if missing",
            "After the imports section, paste the full publish_edh_metrics() function (see Section 6.2)",
            "Find the section in the script where validation results are finalised",
            "Add a call to publish_edh_metrics() passing the job's actual variables",
            "Click Save script",
            "Do NOT run in production — go to Glue → Jobs → Run → select DEV test data → Run job",
            "Watch the run status in Glue → Jobs → Run history",
        ],
        "config": [
            ("Function placement", "After imports, before or inside main function"),
            ("Call placement", "After validation is complete, before job end"),
            ("environment parameter", "'dev' for DEV, 'test' for TEST, 'prod' for PROD"),
            ("boto3 region", "ap-southeast-2 (replace with your AWS region)"),
        ],
        "output": "Script saved. After a successful DEV run, 9 custom metrics appear across 3 CloudWatch namespaces.",
        "validate": "CloudWatch → Metrics → All metrics → Custom Namespaces → verify EDH/Ingestion, EDH/Validation, EDH/Pipeline appear with data points.",
        "tip": "Wrap the publish call in a try/except block so a CloudWatch error does not abort the main pipeline logic.",
    },
    {
        "id": "2.3",
        "title": "Create the /edh/custom/ingestion Log Group",
        "objective": "Creates a dedicated CloudWatch Log group for structured JSON application logs written by the EDH Glue script.",
        "service": "Amazon CloudWatch",
        "time": "3 minutes",
        "steps": [
            "AWS Console → CloudWatch → Log groups (left panel)",
            "Click Create log group button",
            "Log group name field: type /edh/custom/ingestion",
            "Retention — select: 30 days (1 month)",
            "Tags section (optional): click Add tag → Key: Project, Value: EDH",
            "Click Create",
        ],
        "config": [
            ("Log group name", "/edh/custom/ingestion"),
            ("Retention", "30 days"),
            ("Tag: Project", "EDH (optional)"),
        ],
        "output": "/edh/custom/ingestion appears in Log groups list with 30-day retention.",
        "validate": "CloudWatch → Log groups → search /edh → verify the group exists with correct retention.",
        "tip": "If the log group already exists, check that retention is set to 30 days and proceed.",
    },
    {
        "id": "2.4",
        "title": "Verify All Custom Metrics After Test Job Run",
        "objective": "Phase 2 go/no-go checkpoint — confirms PutMetricData calls in the Glue script are successfully reaching CloudWatch.",
        "service": "Amazon CloudWatch",
        "time": "10 minutes",
        "steps": [
            "Open CloudWatch → Metrics → All metrics",
            "On the left panel, scroll to Custom Namespaces section",
            "Verify EDH/Ingestion appears — click it",
            "Click the dimension group: JobName, PipelineName, Environment",
            "Find RecordsProcessed and RecordsRejected → check both checkboxes",
            "Click the Graphed metrics tab at the bottom → verify data points on the graph",
            "Go back → click EDH/Validation → verify ValidationPassCount and ValidationFailCount",
            "Go back → click EDH/Pipeline → verify PipelineExecutionStatus and JobExecutionDuration",
            "For PipelineExecutionStatus: a value of 1 means the last run succeeded",
        ],
        "config": [
            ("Namespaces to verify", "EDH/Ingestion, EDH/Validation, EDH/Pipeline"),
            ("PipelineExecutionStatus success value", "1"),
            ("Data point latency after job", "Within 1–5 minutes"),
        ],
        "output": "Three custom namespaces with data points. PipelineExecutionStatus = 1 for a successful run.",
        "validate": "Every metric in Section 4.2 has at least one data point. No data gaps for the most recent job run.",
        "tip": "If EDH/Ingestion appears but has no data points, set the CloudWatch graph time range to 'Last 3 hours'.",
    },
]

steps_p3 = [
    {
        "id": "3.1",
        "title": "Create SNS Topic and Confirm Email Subscription",
        "objective": "Creates a central notification channel for all EDH operational alerts.",
        "service": "Amazon SNS",
        "time": "5 minutes + email confirmation",
        "steps": [
            "AWS Console → search bar → type SNS → click Simple Notification Service",
            "In left panel, click Topics",
            "Click Create topic button",
            "Type: select Standard (not FIFO)",
            "Name field: type edh-operational-alerts",
            "Display name field: type EDH Ops",
            "Leave all other settings as default",
            "Click Create topic",
            "On the topic page, click Create subscription",
            "Protocol dropdown: select Email",
            "Endpoint field: type your operations team email address",
            "Click Create subscription",
            "Open your email inbox → find the email from no-reply@sns.amazonaws.com",
            "Click the Confirm subscription link in the email",
            "Return to SNS → Subscriptions → verify Status = Confirmed",
        ],
        "config": [
            ("Topic name", "edh-operational-alerts"),
            ("Topic type", "Standard"),
            ("Display name", "EDH Ops"),
            ("Subscription protocol", "Email"),
            ("Subscription status", "Must be Confirmed before alarms work"),
        ],
        "output": "Topic and email subscription confirmed. Topic ARN: arn:aws:sns:{region}:{account-id}:edh-operational-alerts",
        "validate": "SNS → Topics → edh-operational-alerts → Subscriptions tab → Status = Confirmed.",
        "tip": "If the confirmation email does not arrive within 5 minutes, check the spam/junk folder. Resend from SNS → Subscriptions → Request confirmation.",
    },
    {
        "id": "3.2",
        "title": "Create the Three Critical CloudWatch Alarms",
        "objective": "Creates the three most important alarms: Glue task failure, Step Functions execution failure, and overall pipeline execution status.",
        "service": "Amazon CloudWatch",
        "time": "20 minutes",
        "steps": [
            "=== ALARM 1: EDH-GlueJob-TaskFailure ===",
            "CloudWatch → Alarms → All alarms → Create alarm",
            "Click Select metric → AWS Glue → Job Metrics",
            "Find glue.driver.aggregate.numFailedTasks for edh-glue-ingestion-daily → check → Select metric",
            "Statistic: Sum | Period: 5 minutes | Threshold type: Static",
            "Whenever metric is: Greater than | Value: 0 | Datapoints: 1 out of 1",
            "Missing data: Treat as Good | Click Next",
            "Alarm state trigger: In alarm → SNS topic: edh-operational-alerts",
            "Click Next → Alarm name: EDH-GlueJob-TaskFailure → Create alarm",
            "=== ALARM 2: EDH-StepFunctions-ExecutionFailed ===",
            "Create alarm → Select metric → AWS States → Execution Metrics",
            "Find ExecutionsFailed for your EDH state machine → Select metric",
            "Statistic: Sum | Period: 5 min | Threshold: >= 1 | Datapoints: 1 out of 1",
            "Click Next → SNS: edh-operational-alerts → Name: EDH-StepFunctions-ExecutionFailed → Create",
            "=== ALARM 3: EDH-Custom-PipelineFailure ===",
            "Create alarm → Select metric → Custom Namespaces → EDH/Pipeline",
            "Select PipelineExecutionStatus with JobName and Environment dimensions → Select metric",
            "Statistic: Minimum | Period: 5 min | Threshold: Less than 1 | Datapoints: 1 out of 1",
            "Click Next → SNS: edh-operational-alerts → Name: EDH-Custom-PipelineFailure → Create",
        ],
        "config": [
            ("Alarm 1 — Metric", "glue.driver.aggregate.numFailedTasks"),
            ("Alarm 1 — Threshold", "Sum > 0, 1 out of 1 datapoints, 5-minute period"),
            ("Alarm 2 — Metric", "ExecutionsFailed (AWS/States)"),
            ("Alarm 2 — Threshold", "Sum >= 1, 1 out of 1 datapoints, 5-minute period"),
            ("Alarm 3 — Metric", "PipelineExecutionStatus (EDH/Pipeline)"),
            ("Alarm 3 — Threshold", "Minimum < 1, 1 out of 1 datapoints, 5-minute period"),
            ("All alarms — SNS action", "edh-operational-alerts"),
            ("All alarms — Missing data", "Treat as Good"),
        ],
        "output": "Three alarms appear in CloudWatch → Alarms → All alarms with status 'OK' (green) initially.",
        "validate": "Verify all three alarms show OK state after a successful DEV job run. Run a failing job and confirm EDH-GlueJob-TaskFailure transitions to ALARM.",
        "tip": "Initial state is 'Insufficient data' (grey) — normal. Changes to OK once at least one data point exists in the evaluation window.",
    },
    {
        "id": "3.3",
        "title": "Create the EDH Operational CloudWatch Dashboard",
        "objective": "Builds a single-pane operational dashboard showing pipeline health, record counts, execution duration, validation health, and alarm status.",
        "service": "Amazon CloudWatch",
        "time": "25–30 minutes",
        "steps": [
            "CloudWatch → Dashboards (left panel) → Create dashboard",
            "Dashboard name: EDH-Operational-Monitoring → Create dashboard",
            "=== Widget 1: Alarm Status Panel ===",
            "Click + (Add widget) → Alarm status → Next",
            "Title: EDH Pipeline Alarm Status | Select all EDH-* alarms → Create widget",
            "=== Widget 2: Failed Executions (Number) ===",
            "Add widget → Number → Next → Browse → AWS States → Execution Metrics",
            "Select ExecutionsFailed for your state machine",
            "Title: Failed Executions | Period: 1 hour | Statistic: Sum → Create widget",
            "=== Widget 3: Pipeline Execution Duration (Line) ===",
            "Add widget → Line → Next → AWS States → Execution Metrics → ExecutionTime",
            "Statistic: Average | Period: 5 min | Title: Pipeline Execution Duration (ms) → Create widget",
            "=== Widget 4: Records Processed vs Rejected (Line) ===",
            "Add widget → Line → Next → Custom Namespaces → EDH/Ingestion",
            "Select RecordsProcessed AND RecordsRejected (both on same graph)",
            "Title: Records Processed vs Rejected → Create widget",
            "=== Widget 5: Validation Health (Stacked Area) ===",
            "Add widget → Stacked area → Next → EDH/Validation",
            "Select ValidationPassCount and ValidationFailCount",
            "Title: Validation Health → Create widget",
            "=== Widget 6: Glue Heap Memory (Number) ===",
            "Add widget → Number → Next → AWS Glue → Job Metrics",
            "Select glue.ALL.jvm.heap.usage for edh-glue-ingestion-daily",
            "Title: Glue JVM Heap Usage | Statistic: Maximum → Create widget",
            "Click Save dashboard (top right)",
        ],
        "config": [
            ("Dashboard name", "EDH-Operational-Monitoring"),
            ("Widget 1", "Alarm status — all EDH-* alarms"),
            ("Widget 2", "Number — ExecutionsFailed (1-hour Sum)"),
            ("Widget 3", "Line — ExecutionTime (5-min Average)"),
            ("Widget 4", "Line — RecordsProcessed + RecordsRejected"),
            ("Widget 5", "Stacked area — ValidationPassCount + ValidationFailCount"),
            ("Widget 6", "Number — glue.ALL.jvm.heap.usage (Maximum)"),
        ],
        "output": "Dashboard saved with 6 widgets. Alarm status widget shows green for all alarms.",
        "validate": "Open dashboard → run a DEV job → refresh → verify all 6 widgets show data.",
        "tip": "If a widget shows 'Insufficient data', verify dimension values match exactly what your job publishes: CloudWatch → Metrics → Browse.",
    },
    {
        "id": "3.4",
        "title": "Save the Five Logs Insights Queries",
        "objective": "Pre-saves the five EDH investigation queries for one-click access from the Logs Insights Saved queries panel.",
        "service": "Amazon CloudWatch",
        "time": "10 minutes",
        "steps": [
            "CloudWatch → Logs Insights (left panel)",
            "=== Save Query 1: Glue Job Failures ===",
            "Log group(s) selector: click and select /aws-glue/jobs/error",
            "Clear query box → paste the 'Glue job failures' query from Section 5.2",
            "Click Run query to verify it executes without errors",
            "Click Actions dropdown → Save query",
            "Query name: EDH-GlueJobFailures | Folder: type EDH-Monitoring → Save",
            "=== Repeat for all 5 queries from Section 5.2 ===",
            "Save each query with the name and folder from the Config Values table below",
        ],
        "config": [
            ("Folder name (create new)", "EDH-Monitoring"),
            ("Query 1 name", "EDH-GlueJobFailures"),
            ("Query 2 name", "EDH-RecordCounts"),
            ("Query 3 name", "EDH-StepFunctionFailures"),
            ("Query 4 name", "EDH-ValidationFailuresByObject"),
            ("Query 5 name", "EDH-JobDurationTrend"),
        ],
        "output": "5 queries under Logs Insights → Queries → Saved queries → folder EDH-Monitoring.",
        "validate": "Logs Insights → Queries → Saved queries → EDH-Monitoring folder → click each query → verify it loads and runs without syntax errors.",
        "tip": "If /edh/custom/ingestion shows 'No log events found', the Glue script has not yet written to this log group. Run the DEV job after completing Step 2.2 first.",
    },
]

steps_p4 = [
    {
        "id": "4.1",
        "title": "Execute TS-01: Validate the Success Path",
        "objective": "Confirms the entire monitoring stack works for a successful pipeline run.",
        "service": "Multiple",
        "time": "30 minutes",
        "steps": [
            "Glue → ETL Jobs → edh-glue-ingestion-daily → Run (with valid DEV test data)",
            "Step Functions → State machines → your EDH machine → Executions → click running execution",
            "Wait for execution status = SUCCEEDED",
            "S3 → Browse your data lake bucket → verify new Parquet files are present",
            "CloudWatch → Metrics → EDH/Ingestion → RecordsProcessed → verify data point > 0",
            "CloudWatch → Metrics → EDH/Pipeline → PipelineExecutionStatus → verify value = 1",
            "CloudWatch → Dashboards → EDH-Operational-Monitoring → verify all widgets show data",
            "CloudWatch → Alarms → verify no alarm is in ALARM state",
            "Screenshot evidence: Step Functions SUCCEEDED, dashboard populated, PipelineExecutionStatus = 1",
        ],
        "config": [
            ("Expected Step Functions status", "SUCCEEDED"),
            ("Expected PipelineExecutionStatus", "1"),
            ("Expected alarm state", "All alarms remain OK"),
            ("Minimum evidence", "3 screenshots"),
        ],
        "output": "Dashboard green, all alarms OK, all custom metrics populated.",
        "validate": "RecordsProcessed > 0. PipelineExecutionStatus = 1. No alarm fires.",
        "tip": "If RecordsProcessed = 0 but job succeeded: verify that publish_edh_metrics() is passing actual variables, not hardcoded zeros.",
    },
    {
        "id": "4.2",
        "title": "Execute TS-02: Validate the Glue Failure Alarm",
        "objective": "Confirms EDH-GlueJob-TaskFailure transitions to ALARM when the Glue job fails. REVERT the intentional error immediately after capturing evidence.",
        "service": "Multiple",
        "time": "20 minutes",
        "steps": [
            "Glue → ETL Jobs → edh-glue-ingestion-daily → Script tab",
            "At the top of the main() function, add this line:",
            "    raise Exception('EDH-TEST-FAILURE-TS02 — intentional test error')",
            "Click Save script",
            "Glue → Run job",
            "Glue → Jobs → Run history → watch status change to FAILED",
            "Step Functions → Executions → verify execution = FAILED",
            "CloudWatch → Alarms → EDH-GlueJob-TaskFailure → state changes to ALARM (within 5 min)",
            "Check email inbox → confirm SNS notification from 'EDH Ops' was received",
            "CloudWatch → Log groups → /aws-glue/jobs/error → find the exception message",
            "Screenshot: alarm in ALARM state, SNS email, exception in log",
            "=== REVERT — MANDATORY ===",
            "Glue → Script tab → remove the raise Exception line completely → Save script",
            "Run job once more to confirm it succeeds and alarms return to OK",
        ],
        "config": [
            ("Failure code to add (then remove)", "raise Exception('EDH-TEST-FAILURE-TS02')"),
            ("Expected alarm", "EDH-GlueJob-TaskFailure → ALARM within 5 min"),
            ("Expected email", "Notification from EDH Ops SNS topic"),
            ("MANDATORY post-test action", "Remove the raise Exception line and confirm job succeeds"),
        ],
        "output": "Alarm fires, email received, error captured in log. Full failure-to-notification chain validated.",
        "validate": "CloudWatch → Alarms → EDH-GlueJob-TaskFailure → History tab → event 'State changed from OK to ALARM'. After revert: alarm returns to OK.",
        "tip": "If alarm doesn't transition within 5 minutes: CloudWatch → Metrics → AWS Glue → Job Metrics → numFailedTasks → verify data point appeared from the failed run.",
    },
]

def render_phase_steps(phase_steps):
    for step in phase_steps:
        h3(f"Step {step['id']} — {step['title']}")
        p = doc.add_paragraph()
        r = p.add_run(f"Service: {step['service']}   |   Time: {step['time']}")
        r.font.size = Pt(9)
        r.italic = True
        body(f"Objective: {step['objective']}")
        doc.add_paragraph()

        body("Console Steps:")
        for i, action in enumerate(step["steps"]):
            if action.startswith("==="):
                p = doc.add_paragraph()
                r = p.add_run(action.replace("===", "").strip())
                r.bold = True
                r.font.size = Pt(9)
                r.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
                p.paragraph_format.left_indent = Cm(0.5)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.8)
                r1 = p.add_run(f"{i+1}.  ")
                r1.bold = True
                r1.font.size = Pt(9)
                r2 = p.add_run(action)
                r2.font.size = Pt(9)

        doc.add_paragraph()
        body("Configuration Values:")
        cfg_table = doc.add_table(rows=len(step["config"]) + 1, cols=2)
        cfg_table.alignment = WD_TABLE_ALIGNMENT.LEFT
        header_row(cfg_table, ["Field / Setting", "Value to Enter"])
        for ci, (k, v) in enumerate(step["config"]):
            row = cfg_table.rows[ci + 1]
            row.cells[0].text = k
            row.cells[1].text = v
            for cell in row.cells:
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].font.size = Pt(9)
                set_cell_bg(cell, "F5F5F5" if ci % 2 == 1 else "FFFFFF")
        for row in cfg_table.rows:
            row.cells[0].width = Cm(5.5)
            row.cells[1].width = Cm(13)
        set_cell_borders(cfg_table)

        doc.add_paragraph()
        note_box(f"Expected Output: {step['output']}", label="Output")
        note_box(f"How to Validate: {step['validate']}", label="Validate")
        if step.get("tip"):
            note_box(step["tip"], label="Tip / Troubleshoot")
        doc.add_paragraph()

h2("7.1 Phase 1 — Enable Native Metrics")
render_phase_steps(steps_p1)

page_break()
h2("7.2 Phase 2 — IAM Permissions & Custom Metrics")
render_phase_steps(steps_p2)

page_break()
h2("7.3 Phase 3 — SNS Topic, Alarms & Dashboard")
render_phase_steps(steps_p3)

page_break()
h2("7.4 Phase 4 — Testing & PoC Validation")
note_box("Run all tests in DEV only. Capture screenshots of alarms, dashboards, and log results as PoC evidence. Never modify production configurations.", label="Safety")
doc.add_paragraph()
render_phase_steps(steps_p4)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 8. TEST SCENARIOS
# ═══════════════════════════════════════════════════════════════════════════════

h1("8. Test Scenarios")
note_box("All test scenarios must be executed in DEV or a controlled test environment. Never run destructive tests in production.", label="Safety Warning")
doc.add_paragraph()

test_scenarios = [
    {
        "id": "TS-01", "scenario": "Successful Ingestion",
        "trigger": "Run edh-glue-ingestion-daily manually in DEV with valid test data",
        "metric": "RecordsProcessed↑, PipelineExecutionStatus=1, ExecutionsSucceeded↑",
        "log": "/aws-glue/jobs/output: job completion with record count",
        "alarm": "No alarms (success path)",
        "verify": "Step Functions SUCCEEDED; S3 Parquet files created; custom metrics populated",
        "tone": "success",
    },
    {
        "id": "TS-02", "scenario": "Glue Job Failure",
        "trigger": "In DEV: introduce a temporary syntax error in Glue script then re-run",
        "metric": "numFailedTasks>0, ExecutionsFailed↑, PipelineExecutionStatus=0",
        "log": "/aws-glue/jobs/error: stack trace and exception type",
        "alarm": "EDH-GlueJob-TaskFailure, EDH-Custom-PipelineFailure",
        "verify": "CW Alarms shows ALARM state; SNS email received; Step Functions FAILED",
        "tone": "critical",
    },
    {
        "id": "TS-03", "scenario": "Validation Failure",
        "trigger": "In DEV: inject test records that intentionally break validation rules",
        "metric": "ValidationFailCount>0, RecordsRejected increases",
        "log": "/edh/custom/ingestion: validation failure JSON entry",
        "alarm": "EDH-Custom-ValidationFailure",
        "verify": "CW metric ValidationFailCount non-zero; Logs Insights returns failure records; alarm fires",
        "tone": "high",
    },
    {
        "id": "TS-04", "scenario": "Snowflake Orchestration Failure",
        "trigger": "In DEV: suspend Snowflake root task before pipeline run",
        "metric": "ExecutionsFailed↑, DownstreamTaskStatus=0",
        "log": "/aws/states/edh-step-functions: TaskFailed event at Snowflake step",
        "alarm": "EDH-StepFunctions-ExecutionFailed",
        "verify": "Step Functions shows FAILED at Snowflake step; SES notification triggered",
        "tone": "critical",
    },
    {
        "id": "TS-05", "scenario": "Execution Timeout",
        "trigger": "In DEV: set Step Functions timeout to 2 min; run with large dataset",
        "metric": "ExecutionsTimedOut↑, ExecutionTime > threshold",
        "log": "/aws/states/edh-step-functions: States.Timeout event in execution history",
        "alarm": "EDH-StepFunctions-ExecutionTimeout, EDH-StepFunctions-LongRunning",
        "verify": "Step Functions shows TIMED_OUT; both alarms fire within evaluation period",
        "tone": "high",
    },
    {
        "id": "TS-06", "scenario": "Record Count Mismatch",
        "trigger": "In DEV: modify source extract to drop 500+ records; compare counts in Glue",
        "metric": "RecordCountMismatch > 100, SourceRecordCount != TargetRecordCount",
        "log": "/edh/custom/ingestion: source/target mismatch log entry",
        "alarm": "EDH-Custom-RecordMismatch",
        "verify": "CW metric RecordCountMismatch > 100; alarm transitions to ALARM; dashboard divergence visible",
        "tone": "medium",
    },
    {
        "id": "TS-07", "scenario": "Missing S3 Objects",
        "trigger": "In DEV: rename expected S3 Parquet files; trigger downstream validation",
        "metric": "5xxErrors or not-found logged; PipelineExecutionStatus=0",
        "log": "/aws-glue/jobs/error: S3 file not found (FileNotFoundException)",
        "alarm": "EDH-GlueJob-TaskFailure, EDH-S3-ServerErrors",
        "verify": "Glue error log shows S3 path not found; Step Functions FAILED; alarms fire",
        "tone": "critical",
    },
]

tones_map = {"success": "success", "critical": "critical", "high": "high", "medium": "medium", "neutral": "neutral"}

for ts in test_scenarios:
    h3(f"{ts['id']} — {ts['scenario']}")
    add_table(
        ["Field", "Detail"],
        [
            ["How to Trigger",          ts["trigger"]],
            ["Expected Metric Change",  ts["metric"]],
            ["Log Capture Expected",    ts["log"]],
            ["Alarm Expected to Fire",  ts["alarm"]],
            ["Verification Steps",      ts["verify"]],
        ],
        col_widths=[4.5, 14],
        row_tones=[tones_map.get(ts["tone"], "neutral")] * 5
    )
    doc.add_paragraph()

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 9. TASK 1038935 — EVALUATION
# ═══════════════════════════════════════════════════════════════════════════════

h1("9. Task 1038935 — PoC Evaluation")
body("Assessment based on implementation experience and test results. Scored out of 5.")
doc.add_paragraph()

h2("9.1 Evaluation Scores")
add_table(
    ["Criteria", "Score", "Assessment Notes"],
    [
        ["Dashboard Visibility",      "4 / 5", "Rich widget library; native metrics out of the box; custom EDH metrics require Glue script changes"],
        ["Alerting Capabilities",     "4 / 5", "Composite alarms, math expressions, multi-threshold alarms; SNS reliable; ServiceNow needs webhook"],
        ["Operational Usefulness",    "4 / 5", "Covers infrastructure natively; custom metrics add business-level visibility"],
        ["Cross-Pipeline Monitoring", "3 / 5", "Dimension-based filtering works; each new pipeline needs individual dashboard widget setup"],
        ["Multi-Environment Support", "4 / 5", "Account-per-environment or Environment dimension both viable; cross-account adds complexity"],
        ["Implementation Effort",     "3 / 5", "Native: minimal effort; Custom metrics: Glue code + IAM + testing required"],
        ["Operational Overhead",      "4 / 5", "Fully managed; alarm maintenance required; dashboard updates as pipelines grow"],
        ["Technical Limitations",     "3 / 5", "15-month metric retention cap; Snowflake metrics need intermediate layer; PutMetricData costs scale"],
        ["Scalability",               "4 / 5", "Scales to hundreds of metrics; namespace/dimension design is critical for performance"],
        ["Maintainability",           "4 / 5", "Dashboards exportable as JSON; alarm configs repeatable; schema documentation needed"],
        ["Overall",                   "3.7 / 5", "Recommended for EDH-wide adoption. Medium implementation effort. High operational value."],
    ],
    col_widths=[5, 2, 11.5],
    row_tones=["neutral","neutral","neutral","warning","neutral","warning","neutral","warning","neutral","neutral","info"]
)

doc.add_paragraph()
h2("9.2 Technical Limitations")
limitations = [
    ("Metric Retention Cap", "CloudWatch standard metrics are retained for 15 months. For longer historical analysis, export metrics to S3 using CloudWatch Metric Streams or scheduled Lambda exports."),
    ("Snowflake Metrics Require Bridging", "Snowflake downstream task execution results are not natively available in CloudWatch. An intermediate Lambda function (or Glue step) must read from Snowflake metadata tables and publish via PutMetricData."),
    ("PutMetricData Cost at Scale", "CloudWatch charges per custom metric and per PutMetricData API call. At enterprise scale, batch PutMetricData calls (up to 1,000 metrics per call) and design dimensions carefully to avoid exponential metric cardinality."),
    ("Dashboard Maintenance as Pipelines Scale", "Each new EDH pipeline requires manual dashboard widget additions. Consider exporting the dashboard JSON as a template for new pipelines."),
]
for title, detail in limitations:
    p = doc.add_paragraph()
    r1 = p.add_run(f"{title}: ")
    r1.bold = True
    r1.font.size = Pt(10)
    r2 = p.add_run(detail)
    r2.font.size = Pt(10)
    p.paragraph_format.left_indent = Cm(0.3)

doc.add_paragraph()
h2("9.3 Recommendations for EDH-Wide Adoption")
add_table(
    ["Recommendation", "Priority", "Detail"],
    [
        ["Standardise the EDH metric schema",                              "High",   "Define namespace (EDH/), dimension names (JobName, PipelineName, Environment) before scaling"],
        ["Implement PutMetricData from Glue as the default approach",      "High",   "Lowest overhead; reuse Glue scripts; add a shared utility function to the EDH Glue library"],
        ["Enable Glue continuous logging and Step Functions logging",      "High",   "Zero code change; enables Logs Insights queries immediately in all environments"],
        ["Use Environment dimension on all custom metrics",               "High",   "Enables multi-environment monitoring in a single namespace (dev, test, prod)"],
        ["Build a standard EDH alarm template",                           "Medium", "Export alarm JSON from CloudWatch; re-apply per pipeline to avoid manual recreation"],
        ["Create a shared SNS topic per environment",                     "Medium", "One topic per environment (edh-alerts-dev, edh-alerts-prod) reduces subscription overhead"],
        ["Assess ServiceNow integration in a dedicated sprint",           "Low",    "ServiceNow admin, webhook configuration, field mapping, and duplicate-incident suppression required"],
        ["Evaluate Embedded Metric Format (EMF) for future",              "Low",    "EMF in Glue logs automatically creates metrics from structured log lines — reduces PutMetricData calls"],
    ],
    col_widths=[6, 2, 10.5],
    row_tones=["critical","critical","critical","critical","warning","warning","neutral","neutral"]
)

page_break()

# ═══════════════════════════════════════════════════════════════════════════════
# 10. IAM PERMISSIONS REFERENCE
# ═══════════════════════════════════════════════════════════════════════════════

h1("10. IAM Permissions Reference")
body("The following IAM permissions are required to complete the full implementation. Request any missing permissions from your AWS IAM administrator before beginning the relevant phase.")
doc.add_paragraph()

add_table(
    ["IAM Permission", "Required For", "Where to Add", "Phase"],
    [
        ["cloudwatch:PutMetricData",                              "Glue job publishing custom EDH metrics",          "IAM → Roles → Glue job role → Inline policy",                "Phase 2"],
        ["logs:CreateLogGroup\nlogs:CreateLogStream\nlogs:PutLogEvents", "Glue script writing to custom log group", "IAM → Roles → Glue job role → Inline policy",         "Phase 2"],
        ["logs:PutRetentionPolicy",                               "Setting retention on CloudWatch log groups",      "IAM → Users/Roles → CloudWatch admin policy",                "Phase 1"],
        ["logs:CreateLogDelivery\nlogs:PutResourcePolicy",        "Step Functions execution logging",                "IAM → State machine execution role",                         "Phase 1"],
        ["cloudwatch:PutMetricAlarm",                             "Creating CloudWatch alarms",                      "IAM → Users/Roles → CloudWatchFullAccess or custom",          "Phase 3"],
        ["cloudwatch:PutDashboard",                               "Creating CloudWatch dashboards",                  "IAM → Users/Roles → CloudWatchFullAccess or custom",          "Phase 3"],
        ["sns:CreateTopic\nsns:Subscribe\nsns:Publish",           "Creating SNS alert topic and subscriptions",      "IAM → Users/Roles → SNS policy",                             "Phase 3"],
        ["logs:SaveQueryDefinition",                              "Saving Logs Insights queries",                    "IAM → Users/Roles → CloudWatchLogsFullAccess",               "Phase 3"],
    ],
    col_widths=[5.5, 5, 5.5, 2.5],
    row_tones=["info"] * 8
)

doc.add_paragraph()
note_box(
    "If you do not have IAM write permissions, provide the IAM policy JSON from Section 7.2 Step 2.1 "
    "to your AWS administrator and request they attach it to the Glue job's IAM role before you begin Phase 2.",
    label="Note for Teams Without IAM Access"
)

# ═══════════════════════════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════════════════════════

output_path = "/agent/repos/course/EDH_Operational_Monitoring_CloudWatch_PoC.docx"
doc.save(output_path)
print(f"Document saved: {output_path}")
