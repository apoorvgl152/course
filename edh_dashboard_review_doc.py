"""Single Word document: EDH dashboard review, logs, alarms, and ServiceNow."""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
for section in doc.sections:
    section.top_margin = Cm(1.8)
    section.bottom_margin = Cm(1.8)
    section.left_margin = Cm(2.0)
    section.right_margin = Cm(2.0)

NAVY = RGBColor(0x1F, 0x38, 0x64)
BLUE = RGBColor(0x2E, 0x75, 0xB6)


def shade(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_color)
    tcPr.append(shd)


def borders(table):
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            tcBorders = OxmlElement("w:tcBorders")
            for side in ("top", "left", "bottom", "right"):
                border = OxmlElement(f"w:{side}")
                border.set(qn("w:val"), "single")
                border.set(qn("w:sz"), "4")
                border.set(qn("w:space"), "0")
                border.set(qn("w:color"), "CCCCCC")
                tcBorders.append(border)
            tcPr.append(tcBorders)


def add_table(headers, rows, tones=None):
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.LEFT
    tone_bg = {
        "danger": "FDECEA",
        "warning": "FFF4E5",
        "success": "E8F5E9",
        "info": "E8F1FB",
        "neutral": "F7F7F7",
    }
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = header
        run = cell.paragraphs[0].runs[0]
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.LEFT
        shade(cell, "1F3864")
    for r, values in enumerate(rows):
        tone = tones[r] if tones and r < len(tones) else None
        bg = tone_bg.get(tone, "FFFFFF" if r % 2 == 0 else "F7F7F7")
        for c, value in enumerate(values):
            cell = table.rows[r + 1].cells[c]
            cell.text = str(value)
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
            shade(cell, bg)
    borders(table)
    doc.add_paragraph()
    return table


def h1(text):
    p = doc.add_heading(text, level=1)
    p.runs[0].font.color.rgb = NAVY


def h2(text):
    p = doc.add_heading(text, level=2)
    p.runs[0].font.color.rgb = BLUE


def h3(text):
    p = doc.add_heading(text, level=3)
    p.runs[0].font.color.rgb = RGBColor(0x40, 0x40, 0x40)


def body(text):
    p = doc.add_paragraph(text)
    for run in p.runs:
        run.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(8)
    return p


def bullet(text, bold_lead=None):
    p = doc.add_paragraph(style="List Bullet")
    if bold_lead:
        r = p.add_run(bold_lead)
        r.bold = True
        r.font.size = Pt(10.5)
        r2 = p.add_run(text)
        r2.font.size = Pt(10.5)
    else:
        r = p.add_run(text)
        r.font.size = Pt(10.5)
    return p


def note(label, text):
    p = doc.add_paragraph()
    r = p.add_run(f"{label}. ")
    r.bold = True
    r.font.size = Pt(10.5)
    r.font.color.rgb = NAVY
    r2 = p.add_run(text)
    r2.font.size = Pt(10.5)
    p.paragraph_format.space_after = Pt(8)
    return p


def code(text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = "Courier New"
    r.font.size = Pt(8.5)
    p.paragraph_format.left_indent = Cm(0.4)
    p.paragraph_format.space_after = Pt(8)
    return p


def steps(items):
    for i, item in enumerate(items, 1):
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.4)
        r = p.add_run(f"{i}. ")
        r.bold = True
        r.font.size = Pt(10.5)
        r2 = p.add_run(item)
        r2.font.size = Pt(10.5)


# ---------------------------------------------------------------------------
# Cover
# ---------------------------------------------------------------------------
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("\nEDH Operational Monitoring\n")
r.bold = True
r.font.size = Pt(22)
r.font.color.rgb = NAVY

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run("CloudWatch Dashboard Review, Log Groups, Alarms,\nand ServiceNow Integration\n")
r.font.size = Pt(14)
r.font.color.rgb = BLUE

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run(
    "Dashboard: EDH_Operational_Monitoring\n"
    "Glue job: optimized_script_v2\n"
    "State machine: EDH_Metadata_Orchestration_POC\n"
    "S3 bucket: flagstar-edh-testbucket-dev\n"
    "Region: us-east-1 (N. Virginia)\n"
    "Review window: 11 November 2026 to 16 November 2026, UTC, 1 week\n"
)
r.font.size = Pt(11)

note(
    "How to use this document",
    "Part A reviews the dashboard that is already built. Do not rebuild it. "
    "Part B adds log groups and Glue log review. Part C adds alarms, notifications, "
    "and the ServiceNow options. All configuration is done in the AWS Management Console.",
)

doc.add_page_break()

h1("Contents")
for line in [
    "Part A. Dashboard review",
    "1. What is already working",
    "2. How to read the current charts",
    "3. Layout, titles, and which widgets to split",
    "4. Statistics, periods, and reusable dimensions",
    "5. Metrics to add only after they are visible",
    "6. Correlating Step Functions, Glue, and email",
    "7. Monitoring cost",
    "Part B. CloudWatch log groups and Glue execution logs",
    "8. Log groups to configure",
    "9. How to review a Glue execution",
    "10. Logs Insights queries",
    "Part C. Alarms, notifications, and ServiceNow",
    "11. Alarm design",
    "12. Alarm-by-alarm console steps",
    "13. Alert notifications",
    "14. ServiceNow integration options",
    "15. Recommended incident path",
    "16. Implementation order and validation",
]:
    p = doc.add_paragraph(line)
    for run in p.runs:
        run.font.size = Pt(11)

doc.add_page_break()

# ---------------------------------------------------------------------------
# Part A
# ---------------------------------------------------------------------------
h1("Part A. Dashboard review")

h2("1. What is already working")
body(
    "The dashboard EDH_Operational_Monitoring already has 14 widgets. Glue series use "
    "JobRunId = ALL, so later runs of optimized_script_v2 appear without editing the "
    "dashboard. Step Functions series use the state machine ARN, so later executions of "
    "EDH_Metadata_Orchestration_POC appear the same way. S3 storage widgets use the real "
    "bucket flagstar-edh-testbucket-dev. SES Send and Delivery are present as account-level "
    "supporting metrics. Keep this dashboard. The work is to correct scale, statistic, and "
    "order, then add logs and a small set of alarms."
)
body("Readings below come from the one-week screenshots. Treat the numbers as approximate.")

add_table(
    ["What the week shows", "Approximate reading", "Meaning"],
    [
        ["Workflow executions", "Started and succeeded both reach about 4 on 11/10, then stay at 0", "Four successful runs of the state machine. No later workflow runs in this week."],
        ["Workflow duration", "One spike on 11/10, axis top about 3.6 minutes", "Short wall-clock run. The metric is milliseconds. CloudWatch drew the axis in minutes."],
        ["Spark tasks", "Completed tasks and stages move between about 13 and 26 on 11/12–11/14. Failed and killed sit on zero.", "Spark work finished. A failure of 1 would be hard to see on an axis that runs to 26."],
        ["Records and S3 bytes", "Records peak near 201K. S3 read peaks near 15.8 million bytes. S3 write sits near 1.7 million bytes.", "Per-period deltas. The peak is the busiest interval, not the job total."],
        ["Bucket", "Size about 569 MB to 737 MB. Objects about 485 to 666. Tooltip: 496 objects at 2026-11-11 05:00 UTC.", "Daily S3 inventory, not a live ingest counter. Growth matches the Glue window."],
        ["CPU and memory", "Driver CPU near 0.74. Worker CPU near 0. Heap used about 0.5–0.9 GB. Usage ratio near 0.05.", "Worker pressure is low. The ratio is flattened by the byte scale."],
        ["SES", "Sent reaches about 18 near 11/14. Delivered stays around 6–12.", "Account SES activity. It does not line up with the 11/10 workflow runs."],
    ],
    ["info", "success", "success", "warning", "info", "warning", "warning"],
)

note(
    "Glue activity and the workflow are not the same events this week",
    "Glue, Spark, and S3-transfer widgets move on 11/12–11/14. The state machine widgets "
    "move only on 11/10. Those Glue runs were started outside EDH_Metadata_Orchestration_POC, "
    "or the workflow widget points at a different state machine ARN. Compare Glue → ETL jobs → "
    "optimized_script_v2 → Runs with Step Functions → State machines → "
    "EDH_Metadata_Orchestration_POC → Executions before treating the dashboard as one pipeline.",
)

h2("2. How to read the current charts")
body(
    "AWS Glue aggregate metrics are deltas. recordsRead, task counts, bytesRead, S3 filesystem "
    "bytes, shuffle bytes, and elapsedTime report the change since the previous point. With "
    "statistic Sum, each point is the amount in that interval. The area under the line is the "
    "run total. A peak near 201K records means the busiest interval, not the full job count."
)
body(
    "elapsedTime excludes Glue startup. A peak near 7,580 is milliseconds of Spark ETL time in "
    "that interval. It is not the job duration. The state machine ExecutionTime, about 3.6 minutes "
    "on 11/10, is the wall-clock duration of the workflow."
)
body(
    "Heap usage is a ratio from 0 to 1. On the Worker Memory chart it shares an axis with heap "
    "bytes, so a real reading near 0.05 looks like zero. Disk used is in megabytes and is flat "
    "on the zero line of the Resource chart because that axis follows heap bytes."
)
body(
    "BucketSizeBytes and NumberOfObjects are published about once a day. The 05:00 UTC tooltip "
    "is the normal daily storage-metric timestamp. They show that objects landed in the bucket. "
    "They do not show which job wrote them."
)

h2("3. Layout, titles, and which widgets to split")
body("Drag the existing widgets into four bands. Add one text widget at the top. Do not create a second dashboard.")
body(
    "Text widget body: Job optimized_script_v2 · JobRunId ALL · State machine "
    "EDH_Metadata_Orchestration_POC · Bucket flagstar-edh-testbucket-dev · Region us-east-1. "
    "Use a 1-day range while operating. Use 1 week for trend review."
)

add_table(
    ["Band", "Widgets", "Question"],
    [
        ["1. Workflow", "Workflow Execution Monitoring, Workflow Execution Duration", "Did the state machine start, succeed, fail, abort, or time out, and how long did it take?"],
        ["2. Ingestion", "Record Processing, S3 Data Transfer, Source Data Read, S3 Bucket Storage, S3 Object Count", "Did the job read data, write to S3, and did the bucket grow?"],
        ["3. Spark health", "Spark Task Execution, plus one small chart of failed and killed tasks", "Did tasks finish, fail, or get killed?"],
        ["4. Capacity", "CPU, heap usage ratio, heap bytes, disk, shuffle, elapsed time, GC time, SES", "Is the job short of CPU, memory, or disk, and did the account send mail?"],
    ],
)

add_table(
    ["Current title", "Clearer title"],
    [
        ["Resource Monitoring", "Glue disk used (MB)"],
        ["S3 Data Transfer Monitoring", "Glue S3 bytes read and written per period"],
        ["Record Processing Monitoring", "Glue records read per period"],
        ["Spark Task Execution Monitoring", "Spark tasks and stages per period"],
        ["Glue Execution Performance Monitoring", "Spark ETL elapsed time and GC time per period"],
        ["Source Data Read Monitoring", "Source bytes read per period, all sources"],
        ["Worker Memory Monitoring", "Worker 1 JVM heap used (bytes)"],
        ["SES Email Activity Monitoring", "Account SES send and delivery, not EDH-only"],
    ],
)

body("Keep these widgets together. They already share a unit.")
bullet("S3 bytes read with S3 bytes written.")
bullet("Shuffle bytes written with shuffle local bytes read.")
bullet("Driver CPU with overall worker CPU. Driver near 0.74 and worker near 0 means this job is not worker-CPU bound.")
bullet("Completed tasks with completed stages.")
bullet("SES Send with SES Delivery, both on statistic Sum.")
bullet("Bucket size beside object count, as two widgets. Do not put bytes and object count on one axis.")

body("Split these widgets. The current scale hides the signal.")
bullet("Worker memory: keep glue.1.jvm.heap.used. Graph glue.1.jvm.heap.usage on its own widget with axis label Ratio (0–1).", None)
bullet("Resource monitoring: graph glue.driver.BlockManager.disk.diskSpaceUsed_MB alone, in megabytes.", None)
bullet("Spark tasks: keep completed tasks and stages. Add a small widget with only glue.driver.aggregate.numFailedTasks and glue.driver.aggregate.numKilledTasks.", None)
bullet("Elapsed time and GC time: keep them together at the bottom as a troubleshooting widget. Do not read the peak as job duration.", None)

body("Leave these separate.")
bullet("Source bytes read and S3 filesystem bytes. bytesRead includes every source and updates when a Spark task finishes. S3 filesystem bytes are S3 reads since the previous report, including extra S3 reads.")
bullet("Spark elapsed time and state-machine duration.")
bullet("Daily bucket size and per-period S3 bytes written.")

note(
    "Check the heap metric name",
    "glue.driver.jvm.heap.used is the driver. glue.1.jvm.heap.used is executor 1. "
    "glue.ALL.jvm.heap.used is all executors. Open each widget, Graphed metrics, and read the "
    "metric name. Prefer glue.ALL.jvm.heap.usage for the capacity band once that series is visible. "
    "Executor 1 alone is a weak operational view if the job later uses more workers.",
)

h2("4. Statistics, periods, and reusable dimensions")
body(
    "Glue records, task counts, bytes, shuffle, S3 filesystem bytes, and elapsed time are deltas. "
    "Use Sum. Heap, disk, and CPU are gauges. Use Average. Sum on a gauge, or Average on a delta, "
    "changes the meaning of the line."
)

add_table(
    ["Series", "Statistic", "Period", "What one point means"],
    [
        ["records, tasks, stages, bytes, shuffle", "Sum", "1 minute during a run; 1 hour on a week view", "Amount in that bucket. Add the points for the run total."],
        ["elapsedTime and GC time", "Sum", "1 minute", "Milliseconds of ETL time in that bucket. Not wall-clock duration."],
        ["heap, disk, CPU", "Average", "1 minute", "Level at that time. Maximum shows the peak."],
        ["ExecutionsStarted, Succeeded, Failed, Aborted, TimedOut", "Sum", "5 minutes or 1 hour", "Execution count. Do not use SampleCount."],
        ["ExecutionTime", "Maximum", "5 minutes", "Longest state-machine duration in the bucket, in milliseconds."],
        ["BucketSizeBytes, NumberOfObjects", "Average", "1 day", "Daily inventory."],
        ["SES Send and Delivery", "Sum", "1 hour or 1 day", "Account messages. Both rows need Sum."],
    ],
)

note(
    "SampleCount on ExecutionsStarted",
    "Step Functions emits ExecutionsStarted as 1 when a run starts and as 0 when it ends. "
    "Sum still equals the number of starts. SampleCount doubles it. The chart peak near 4 "
    "is consistent with Sum. Confirm SampleCount is not selected.",
)

body("Edit path: CloudWatch → Dashboards → EDH_Operational_Monitoring → widget menu → Edit → Graphed metrics. Set Statistic and Period. Then Options → left Y-axis label: Bytes, Count, Milliseconds, Megabytes, or Ratio (0–1). Update widget. Save dashboard.")
body("Keep these dimensions so future runs appear automatically.")
bullet("Glue: JobName = optimized_script_v2, JobRunId = ALL, Type = count for aggregate counters and Type = gauge for heap, disk, and CPU.")
bullet("Step Functions: StateMachineArn of EDH_Metadata_Orchestration_POC. Do not select an execution ARN.")
bullet("S3 size: BucketName = flagstar-edh-testbucket-dev and StorageType = StandardStorage.")
bullet("S3 object count: same bucket and StorageType = AllStorageTypes.")
body(
    "A different Glue job name will not appear on these widgets. Two overlapping runs of "
    "optimized_script_v2 are added together on JobRunId ALL."
)

h2("5. Metrics to add only after they are visible")
body("CloudWatch → Metrics → All metrics. Search before adding a widget. If the series is absent, do not add it.")

add_table(
    ["Metric", "Where it comes from", "Add it when"],
    [
        ["glue.driver.ExecutorAllocationManager.executors.numberAllExecutors", "Existing Glue job metrics. Type = gauge. Statistic Average.", "You can see it under the Glue namespace for this job."],
        ["glue.driver.ExecutorAllocationManager.executors.numberMaxNeededExecutors", "Existing Glue job metrics. Statistic Maximum.", "Same widget. A gap versus running executors means Spark was waiting for capacity."],
        ["ExecutionsFailed and ExecutionsTimedOut", "AWS/States. Statistic Sum. StateMachineArn.", "The workflow legend is scrolled or the series were never added. Enlarge the widget first."],
        ["glue.succeed.ALL and glue.error.ALL", "Glue observability metrics. ObservabilityGroup = error. Statistic Sum.", "The series already exist. Do not turn observability on only to create them."],
        ["EventBridge Invocations and FailedInvocations", "The schedule rule that targets this state machine.", "You have confirmed the rule name. These are standard metrics."],
    ],
    ["info", "info", "success", "warning", "neutral"],
)

note(
    "Job success is not a classic Glue metric",
    "numFailedTasks counts Spark tasks. It stays at zero when the script fails before tasks start. "
    "Job status is on Glue → Runs, and on the state machine execution when the state machine started the job. "
    "glue.error.ALL exists only when observability metrics are enabled.",
)
body(
    "Leave out S3 request metrics, SES reputation, and a custom records-written metric. "
    "The standard Glue aggregate set has recordsRead. It does not have a matching records-written aggregate. "
    "S3 bytes written plus daily object growth cover ingestion for this PoC."
)

h2("6. Correlating Step Functions, Glue, and email")
add_table(
    ["Link", "Where it already exists"],
    [
        ["Workflow execution to Glue run", "Step Functions execution history. The Glue StartJobRun output contains the Glue job run ID."],
        ["Glue run to a dashboard point", "JobRunId ALL includes every run of optimized_script_v2, including a run started from the Glue console."],
        ["Workflow failure to the SES email", "The execution shows the SES task. The SES widget stays account-wide until the send uses a configuration set."],
        ["Script failure before Spark starts", "Glue → optimized_script_v2 → Runs. Failed-task count can remain zero."],
        ["Schedule fired", "EventBridge → the rule that targets this state machine → Monitoring."],
    ],
)

h2("7. Monitoring cost")
add_table(
    ["Setting", "What to do"],
    [
        ["Glue Job metrics", "Leave on while this dashboard is in use. These are custom metrics, including executor series that are not on a chart."],
        ["Glue observability metrics", "Turn off if no widget uses them. They are a second custom-metric family and are often enabled by default on Glue 4.0 and later in Glue Studio."],
        ["Continuous logging", "Turn off when you are not debugging a live run. The volume charge is log group /aws-glue/jobs/logs-v2."],
        ["Dashboard refresh", "Refresh manually. Glue points arrive about once a minute."],
        ["S3 storage metrics", "Keep. Do not enable S3 request metrics for this PoC."],
        ["Alarms", "Add the five alarms in Part C. Each standard alarm has a monthly charge. Do not add a large alarm pack."],
    ],
)
body(
    "To pause without deleting the dashboard: Glue → ETL jobs → optimized_script_v2 → Job details → "
    "Advanced properties. Turn off Job metrics, Job observability metrics, and Continuous logging. Save. "
    "Confirm Job parameters does not list --enable-metrics. A value of false still enables that parameter. "
    "The parameter has to be absent. Historical points and the dashboard remain. Turn Job metrics back on "
    "before the next run you want on the charts. In Cost Explorer, filter CloudWatch usage types "
    "MetricMonitorUsage and DataProcessing-Bytes."
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# Part B logs
# ---------------------------------------------------------------------------
h1("Part B. CloudWatch log groups and Glue execution logs")

h2("8. Log groups to configure")
body(
    "Logs explain why a run failed. The dashboard shows that something moved. "
    "The log group shows the exception, the executor, and the job run ID. "
    "Create or confirm the groups below in us-east-1. Set retention so log storage does not grow without a limit."
)

add_table(
    ["Log group", "What it contains", "When it receives logs", "Retention to set"],
    [
        ["/aws-glue/jobs/output", "Glue driver and executor standard output for optimized_script_v2.", "At the end of the job, and during the job if continuous logging is on.", "30 days"],
        ["/aws-glue/jobs/error", "Exceptions and error output.", "Same as output. This is the first place to look after a failure.", "30 days"],
        ["/aws-glue/jobs/logs-v2", "Near-real-time driver and executor streams.", "Only while Continuous logging is enabled on the job.", "14 days"],
        ["/aws/vendedlogs/states/EDH_Metadata_Orchestration_POC", "State transitions, task failures, and the Glue job run ID passed between states.", "After logging is enabled on the state machine.", "30 days"],
    ],
)

h3("8.1 Confirm Glue logging")
steps([
    "Open the AWS console in us-east-1. Go to AWS Glue → ETL jobs → optimized_script_v2.",
    "Open the Job details tab. Scroll to Advanced properties.",
    "Confirm whether Continuous logging is on. For a failure investigation, turn it on. For an idle PoC week, it can stay off to limit ingestion cost. Output and error log groups still receive logs at the end of a job when the job itself writes logs.",
    "If the console offers to attach a CloudWatch Logs policy to the Glue role, accept it. The role needs logs:CreateLogGroup, logs:CreateLogStream, and logs:PutLogEvents. If Save fails with an access message, request that policy from the AWS administrator.",
    "Save the job.",
    "Open CloudWatch → Log groups. Search /aws-glue/jobs. Confirm output, error, and, if continuous logging is on, logs-v2.",
    "Select each group → Actions → Edit retention setting. Choose 30 days for output and error, and 14 days for logs-v2. Save.",
])
note(
    "Permission if retention is greyed out",
    "The console user needs logs:PutRetentionPolicy. Request it from the AWS administrator. "
    "The dashboard and existing log events stay in place while that permission is requested.",
)

h3("8.2 Enable Step Functions logging")
body(
    "Let the Step Functions console create the log group. That attaches the resource policy the "
    "service needs. A log group created by hand often rejects delivery until that policy exists."
)
steps([
    "Open Step Functions → State machines → EDH_Metadata_Orchestration_POC.",
    "Choose Edit.",
    "Open Logging. Set Log level to ERROR for normal operation. Use ALL only while you are reproducing one failure, then return it to ERROR. ALL records every state transition and increases log volume.",
    "Choose Create new log group if one is not already selected. Use the name /aws/vendedlogs/states/EDH_Metadata_Orchestration_POC.",
    "Turn on Include execution data while testing in this dev account so the Glue job run ID and the error cause are visible. Execution data can contain job input. Keep this off if the input includes sensitive values, and use ERROR level without execution data.",
    "Leave X-Ray tracing off for this PoC.",
    "Save. If Save reports an IAM error, the state machine role needs logs:CreateLogDelivery, logs:GetLogDelivery, logs:UpdateLogDelivery, logs:DeleteLogDelivery, logs:ListLogDeliveries, logs:PutResourcePolicy, logs:DescribeResourcePolicies, and logs:DescribeLogGroups. Give that list to the AWS administrator.",
    "Run one successful execution. CloudWatch → Log groups → the new group should show one log stream for that execution.",
])

h3("8.3 What not to create")
body(
    "Do not create a second custom application log group for this pass. The Glue output, Glue error, "
    "and Step Functions log groups cover execution review. A custom log group is useful later if the "
    "script writes structured record counts. It is not required to alarm on job failure, delay, or heap."
)

h2("9. How to review a Glue execution")
body("Use the job run ID as the join key. It looks like jr_ followed by an identifier.")
steps([
    "AWS Glue → ETL jobs → optimized_script_v2 → Runs. Note the run ID, start time, status, and error message shown on the run.",
    "If the run was started by the state machine, open Step Functions → EDH_Metadata_Orchestration_POC → Executions → the execution at the same time. Open the Glue task. The output contains the same job run ID. The failed state shows the error and cause.",
    "CloudWatch → Log groups → /aws-glue/jobs/error. In the log-stream search box, paste the job run ID.",
    "Open the driver stream first. Search inside the stream for Exception, Error, or Caused by. Then open executor streams if the driver points at a lost or killed executor.",
    "Repeat the search in /aws-glue/jobs/output for the lines just before the exception. Output often has the Spark stage number.",
    "If continuous logging is on, repeat the search in /aws-glue/jobs/logs-v2. Those streams update while the job is still running.",
    "Return to the dashboard and set the time range to that run. Confirm whether numFailedTasks or numKilledTasks moved, and whether the state machine shows ExecutionsFailed.",
])

add_table(
    ["What you see in the log", "What it means for this pipeline", "Where the dashboard should move"],
    [
        ["AnalysisException, SyntaxError, or a Python traceback before any stage", "The script failed before Spark tasks ran.", "numFailedTasks can stay at 0. The state machine should show ExecutionsFailed if it started the job."],
        ["ExecutorLostFailure or Container killed by YARN", "An executor ran out of memory or disk.", "numKilledTasks and heap or disk should move. Heap usage near 0.05 this week does not show this condition."],
        ["FileNotFoundException or NoSuchKey for an S3 path", "The expected Parquet object was missing.", "The Glue error log is the evidence. BucketSizeBytes will not drop just because one key is missing."],
        ["States.Timeout or a Glue task timeout in the state log", "The workflow exceeded its configured timeout.", "ExecutionsTimedOut should increment. ExecutionTime is only published when the execution closes."],
        ["AccessDenied on cloudwatch:PutMetricData or logs:PutLogEvents", "The Glue role can run the job and still fail while publishing metrics or logs.", "Dashboard gaps for that run. Fix the role before treating the chart as complete."],
    ],
    ["warning", "danger", "warning", "danger", "info"],
)

h2("10. Logs Insights queries")
body(
    "CloudWatch → Logs → Logs Insights. Select the log group, set the time range to the run, "
    "paste the query, and choose Run query. Save each query with Actions → Save query, in a folder named EDH-Monitoring."
)

h3("Glue errors for one run")
body("Log group: /aws-glue/jobs/error. Replace the run ID.")
code(
    "fields @timestamp, @logStream, @message\n"
    "| filter @logStream like /jr_REPLACE_RUN_ID/\n"
    "| filter @message like /ERROR|Exception|Caused by|FAILED/\n"
    "| sort @timestamp asc\n"
    "| limit 100"
)

h3("Latest Glue errors across runs")
body("Log group: /aws-glue/jobs/error.")
code(
    "fields @timestamp, @logStream, @message\n"
    "| filter @message like /ERROR|Exception|FAILED/\n"
    "| sort @timestamp desc\n"
    "| limit 50"
)

h3("Step Functions task failures")
body("Log group: /aws/vendedlogs/states/EDH_Metadata_Orchestration_POC. This returns rows only after section 8.2 is saved and a run has failed or been logged.")
code(
    "fields @timestamp, type, details.error, details.cause\n"
    "| filter type = \"ExecutionFailed\" or type = \"TaskFailed\" or type = \"ExecutionTimedOut\"\n"
    "| sort @timestamp desc\n"
    "| limit 50"
)

h3("Find the Glue run ID inside a workflow log")
body("Same Step Functions log group.")
code(
    "fields @timestamp, @message\n"
    "| filter @message like /jr_/\n"
    "| sort @timestamp desc\n"
    "| limit 20"
)

note(
    "If a query returns no rows",
    "Confirm the time range covers the run, the log group name matches the account, and the job has "
    "finished at least once after logging was enabled. An empty logs-v2 group means continuous logging "
    "is off. Use /aws-glue/jobs/error instead.",
)

doc.add_page_break()

# ---------------------------------------------------------------------------
# Part C alarms
# ---------------------------------------------------------------------------
h1("Part C. Alarms, notifications, and ServiceNow")

h2("11. Alarm design")
body(
    "An alarm watches one metric and changes state. It does not open a ServiceNow incident by itself. "
    "The path is: CloudWatch evaluates the metric, the alarm enters ALARM, Amazon SNS sends the message, "
    "and ServiceNow creates or updates an incident only if a subscription is connected."
)
body(
    "The state machine already sends an SES email when an execution fails. Keep that email. "
    "Add alarms for the conditions the email does not cover well: Spark task failure, a run that "
    "finishes slowly, a timeout, and heap pressure. Do not also create a second incident for the "
    "same execution failure. Section 15 chooses one incident path."
)

add_table(
    ["Alarm name", "Metric", "Threshold", "What it catches"],
    [
        ["EDH-Glue-FailedTasks", "glue.driver.aggregate.numFailedTasks", "Sum > 0 for 1 minute", "Spark tasks failed inside optimized_script_v2."],
        ["EDH-Glue-KilledTasks", "glue.driver.aggregate.numKilledTasks", "Sum > 0 for 1 minute", "Tasks killed, often memory, disk, or skew."],
        ["EDH-SF-ExecutionsFailed", "ExecutionsFailed", "Sum >= 1 for 5 minutes", "The state machine execution failed."],
        ["EDH-SF-ExecutionsTimedOut", "ExecutionsTimedOut", "Sum >= 1 for 5 minutes", "The execution hit its configured timeout."],
        ["EDH-SF-ExecutionDelay", "ExecutionTime", "Maximum > 900000 ms (15 minutes)", "A finished execution took longer than about four times the 3.6 minute run seen on 11/10."],
        ["EDH-Glue-HeapUsageHigh", "glue.ALL.jvm.heap.usage, or glue.1.jvm.heap.usage if ALL is absent", "Average > 0.85 for 3 of 3 five-minute points", "Heap pressure before an out-of-memory failure. Current week is near 0.05, so this stays OK."],
    ],
    ["danger", "danger", "danger", "warning", "warning", "info"],
)

note(
    "Missing data must be Not breaching",
    "These metrics are quiet when no job is running. If missing data is treated as breaching, "
    "the alarm stays in ALARM after the job ends. On every alarm, set Missing data treatment to "
    "Treat missing data as good (not breaching).",
)
note(
    "15 minutes is a starting point, not an SLA",
    "One observed workflow lasted about 3.6 minutes. 900000 milliseconds is 15 minutes. "
    "Change it after several real runs. Do not lower the state machine timeout to 15 minutes "
    "from this document. Check the current timeout first. A timeout fails the execution. "
    "The delay alarm only marks a finished execution that was slow.",
)
note(
    "ExecutionTime does not warn while a job is still running",
    "AWS publishes ExecutionTime when the execution closes. A run that is stuck past 15 minutes "
    "and has not finished will not move this alarm yet. The timeout on the state machine, plus "
    "EDH-SF-ExecutionsTimedOut, covers a run that never closes. View the current TimeoutSeconds "
    "on the state machine before changing it.",
)

h2("12. Alarm-by-alarm console steps")
body("Create the SNS topic in section 13 first, confirm the email subscription, then create the alarms. Region: us-east-1.")

h3("12.1 EDH-Glue-FailedTasks")
steps([
    "CloudWatch → Alarms → All alarms → Create alarm.",
    "Select metric → Glue. Search numFailedTasks.",
    "Select the series for JobName = optimized_script_v2, JobRunId = ALL, and Type = count. If several rows appear, choose the row whose Type is count.",
    "Statistic: Sum. Period: 1 minute.",
    "Threshold type: Static. Whenever numFailedTasks is Greater than 0.",
    "Datapoints to alarm: 1 out of 1.",
    "Missing data treatment: Treat missing data as good (not breaching).",
    "Next. Alarm state trigger: In alarm. Also select OK, so a recovery notice is sent. Choose the SNS topic edh-operational-alerts.",
    "Alarm name: EDH-Glue-FailedTasks. Description: Spark tasks failed for optimized_script_v2.",
    "Create alarm.",
])

h3("12.2 EDH-Glue-KilledTasks")
body("Repeat section 12.1 with metric glue.driver.aggregate.numKilledTasks, the same dimensions, Sum, 1 minute, Greater than 0, missing data not breaching. Name: EDH-Glue-KilledTasks.")

h3("12.3 EDH-SF-ExecutionsFailed")
steps([
    "Create alarm → Select metric → States → Execution Metrics.",
    "Select ExecutionsFailed for EDH_Metadata_Orchestration_POC. The dimension StateMachineArn fills in from the console. Do not type an account ID by hand.",
    "Statistic: Sum. Period: 5 minutes. Threshold: Greater than or equal to 1. Datapoints: 1 out of 1.",
    "Missing data: Treat missing data as good (not breaching).",
    "Notify the same SNS topic on ALARM and OK.",
    "Name: EDH-SF-ExecutionsFailed. Create alarm.",
])

h3("12.4 EDH-SF-ExecutionsTimedOut")
body("Repeat section 12.3 with metric ExecutionsTimedOut. Name: EDH-SF-ExecutionsTimedOut. Same period, statistic, threshold, missing-data setting, and SNS topic.")

h3("12.5 EDH-SF-ExecutionDelay")
steps([
    "Create alarm → States → Execution Metrics → ExecutionTime for EDH_Metadata_Orchestration_POC.",
    "Statistic: Maximum. Period: 5 minutes.",
    "Threshold: Greater than 900000. The unit is milliseconds. 900000 ms is 15 minutes.",
    "Datapoints: 1 out of 1. Missing data: not breaching.",
    "Notify the same SNS topic on ALARM and OK.",
    "Name: EDH-SF-ExecutionDelay. Description: Finished execution exceeded 15 minutes. Baseline on 11/10 was about 3.6 minutes.",
    "Create alarm.",
])

h3("12.6 EDH-Glue-HeapUsageHigh")
steps([
    "CloudWatch → Metrics → Glue. Search heap.usage.",
    "Prefer glue.ALL.jvm.heap.usage with JobName = optimized_script_v2, JobRunId = ALL, Type = gauge. If that series is not listed, use glue.1.jvm.heap.usage and record that the alarm covers executor 1 only.",
    "Statistic: Average. Period: 5 minutes.",
    "Threshold: Greater than 0.85. Datapoints to alarm: 3 out of 3, so one short spike does not page.",
    "Missing data: not breaching. SNS topic on ALARM and OK.",
    "Name: EDH-Glue-HeapUsageHigh. Create alarm.",
])

h3("12.7 Put the alarms on the existing dashboard")
steps([
    "CloudWatch → Dashboards → EDH_Operational_Monitoring.",
    "Add widget → Alarm status.",
    "Select the six EDH alarms. Title: EDH alarm status.",
    "Drag this widget into band 1, beside the workflow charts.",
    "Save dashboard.",
])
body(
    "Do not add an alarm on BucketSizeBytes, SES Send, or Spark elapsed time. Bucket size is a daily "
    "inventory and will not detect a failed job. SES Send is account-wide. Elapsed time is a per-minute "
    "delta, so a threshold on one point does not mean the job ran too long."
)

h2("13. Alert notifications")
body(
    "Use one standard SNS topic for every EDH alarm. A separate topic per alarm makes ServiceNow "
    "correlation harder and produces duplicate subscriptions."
)

h3("13.1 Create and confirm the topic")
steps([
    "Amazon SNS → Topics → Create topic.",
    "Type: Standard. Name: edh-operational-alerts. Display name: EDH Ops.",
    "Leave encryption and delivery policy at the defaults for this PoC. Create topic.",
    "Create subscription. Protocol: Email. Endpoint: the operations mailbox that should see PoC alerts.",
    "Create a second email subscription for yourself while testing, if that mailbox is different.",
    "Open the confirmation message from no-reply@sns.amazonaws.com and confirm each subscription. Check junk mail if it does not arrive within 5 minutes.",
    "SNS → Subscriptions. Status must be Confirmed. PendingConfirmation does not deliver alarms.",
])

h3("13.2 What each message means")
add_table(
    ["State", "Who acts", "What it is"],
    [
        ["OK", "CloudWatch", "The metric is inside the threshold, or there is no data and missing data is not breaching."],
        ["ALARM", "CloudWatch, then SNS", "The threshold was crossed for the required datapoints. SNS sends email. This is the alert. It is not yet an incident."],
        ["INSUFFICIENT_DATA", "CloudWatch", "The alarm has not seen enough points. New alarms start here. It should change to OK after a quiet period because missing data is not breaching."],
    ],
)

h3("13.3 Test the notification without touching production data")
body("Use the dev job and the dev bucket only.")
steps([
    "Glue → optimized_script_v2 → Script. At the first line of the job body, add a temporary line: raise Exception('EDH-TEST-ALARM'). Save.",
    "Start the job from EDH_Metadata_Orchestration_POC, not from a separate console run, so the workflow alarm and the Glue log describe the same event.",
    "Within about 5 minutes, EDH-SF-ExecutionsFailed should enter ALARM. The email from EDH Ops should arrive. The Glue error log should contain EDH-TEST-ALARM.",
    "EDH-Glue-FailedTasks may stay OK. A Python exception before Spark tasks does not increment numFailedTasks. That is expected and is why the state-machine alarm exists.",
    "Remove the raise line. Save the script. Run the state machine once more. The failed-execution alarm should return to OK and send an OK message.",
    "Screenshot the ALARM state, the email, the error log stream, and the OK recovery for the PoC evidence.",
])
note(
    "Heap and delay alarms",
    "Do not force a 15-minute run or fill the heap to test those alarms in this pass. "
    "Confirm each alarm exists, the statistic and missing-data setting are correct, and the state is OK or INSUFFICIENT_DATA. "
    "A later slow run or a real memory rise will exercise them.",
)

h2("14. ServiceNow integration options")
body(
    "The existing SES email from the state machine already notifies people when an execution fails. "
    "ServiceNow integration is a separate decision. Confirm with the ServiceNow owner whether that "
    "mailbox already opens an incident. The right option depends on that answer. Do not build a "
    "second incident path until the owner answers."
)

add_table(
    ["Option", "How it works", "Effort", "Fit for this PoC"],
    [
        ["A. Existing mailbox", "SNS or the current SES email goes to an address that ServiceNow already converts into an incident.", "Low, if that mailbox exists.", "Best first step. No new AWS service. Incident quality depends on the email parser."],
        ["B. SNS HTTPS subscription", "The alarm topic sends the CloudWatch JSON to a ServiceNow Scripted REST API or Event Management endpoint. ServiceNow confirms the subscription.", "Medium. Needs a ServiceNow admin and an endpoint URL.", "Best structured path. Alarm name, state, and reason become fields. OK can resolve the incident."],
        ["C. IntegrationHub AWS spoke", "ServiceNow pulls or receives AWS events through its AWS integration.", "Medium to high. Owned by the ServiceNow team.", "Use this if the enterprise standard is IntegrationHub. Do not invent a parallel webhook."],
        ["D. Lambda between SNS and ServiceNow", "A function reshapes the alarm into the ServiceNow incident API.", "High. Code, credentials, and retries.", "Not recommended for this PoC. Option B covers the same need."],
    ],
    ["success", "info", "neutral", "warning"],
)

h3("14.1 Fields ServiceNow should map")
body(
    "A CloudWatch alarm notification is JSON. The ServiceNow transform should read these fields and "
    "should not open a new incident when NewStateValue is OK."
)
add_table(
    ["JSON field", "Use in the incident"],
    [
        ["AlarmName", "Correlation key, together with the job or state machine name. Example: EDH-SF-ExecutionsFailed."],
        ["NewStateValue", "ALARM opens or updates the incident. OK resolves it or adds a work note."],
        ["NewStateReason", "Short description of the threshold that fired."],
        ["StateChangeTime", "Incident opened time."],
        ["Region", "us-east-1."],
        ["Trigger.MetricName and Trigger.Dimensions", "Glue job name or state machine ARN. This separates EDH from other alarms on the same topic."],
    ],
)

h3("14.2 Console steps for option B, after ServiceNow provides the URL")
steps([
    "Ask the ServiceNow administrator for the inbound Event Management or Scripted REST URL that accepts CloudWatch alarm notifications. Do not guess the URL.",
    "SNS → Topics → edh-operational-alerts → Create subscription.",
    "Protocol: HTTPS. Endpoint: the URL supplied by ServiceNow. Enable raw message delivery only if the ServiceNow owner asks for it. Otherwise leave it off so ServiceNow receives the SNS envelope and can confirm the subscription.",
    "Create subscription. Status will be PendingConfirmation.",
    "The ServiceNow endpoint must call the SubscribeURL from the confirmation message. The ServiceNow owner does this. AWS will not mark the subscription Confirmed until that call succeeds.",
    "After status is Confirmed, put one alarm into ALARM with the test in section 13.3. Confirm that ServiceNow opens one incident and that the following OK message updates or closes that same incident.",
])

h3("14.3 Duplicate incidents to avoid")
bullet("The state machine SES email and EDH-SF-ExecutionsFailed both describe the same failed execution. If both create incidents, one failure becomes two tickets.")
bullet("EDH-Glue-FailedTasks and EDH-SF-ExecutionsFailed can both fire when Spark tasks fail inside a workflow run. Correlate them in ServiceNow on the execution time and the job name, or page on the state-machine alarm and keep the Glue task alarm as email only.")
bullet("Account-wide SES Send must not open an incident. It includes mail that is not from this pipeline. Do not alarm on it.")
bullet("An OK notification must not open a new incident.")

h2("15. Recommended incident path")
body("Use this split so monitoring, alerting, and incident creation stay distinct.")

add_table(
    ["Condition", "Monitoring", "Alert", "Incident"],
    [
        ["State machine execution failed", "ExecutionsFailed on the dashboard", "EDH-SF-ExecutionsFailed sends SNS email", "One ServiceNow incident, from either the existing SES process or option B. Not both."],
        ["Execution timed out", "ExecutionsTimedOut", "EDH-SF-ExecutionsTimedOut", "Same incident path as a failed execution."],
        ["Finished run slower than 15 minutes", "ExecutionTime widget", "EDH-SF-ExecutionDelay email", "Email for the PoC. Open an incident only if the ServiceNow owner wants slow-run tickets."],
        ["Spark tasks failed or were killed", "Task widgets and Glue error log", "EDH-Glue-FailedTasks and EDH-Glue-KilledTasks email", "Email. Let the execution-failed incident be the ticket when the workflow started the job."],
        ["Heap above 0.85 for 15 minutes", "Heap usage widget", "EDH-Glue-HeapUsageHigh email", "Email during the PoC. This is an early warning, not an outage by itself."],
        ["Schedule did not fire", "EventBridge FailedInvocations, after the rule name is confirmed", "Add that alarm later", "Incident only after the rule is verified."],
    ],
    ["danger", "danger", "warning", "warning", "info", "neutral"],
)

body(
    "Recommendation for this PoC: confirm option A with the ServiceNow owner first. If the existing "
    "failure mailbox already opens one incident, subscribe edh-operational-alerts to that same mailbox "
    "and do not add an HTTPS subscription yet. Choose option B when the owner wants alarm name and "
    "OK recovery as structured fields. Defer option C unless it is already the company standard. "
    "Do not build option D."
)

h2("16. Implementation order and validation")
add_table(
    ["Step", "Action", "Done when"],
    [
        ["1", "Compare Glue Runs with state machine executions for 11/10–11/16.", "You know which Glue runs were outside the state machine."],
        ["2", "Fix widget statistics, Y-axis labels, and the heap and disk split on the existing dashboard.", "Worker usage has its own 0–1 axis. Disk has its own megabyte axis. Byte charts are labeled Bytes."],
        ["3", "Set retention on /aws-glue/jobs/output and /aws-glue/jobs/error to 30 days.", "Retention column shows 30 days."],
        ["4", "Enable Step Functions logging at ERROR and create the vended log group.", "One successful execution writes a log stream."],
        ["5", "Save the four Logs Insights queries.", "Each query runs, or returns no rows with an explanation."],
        ["6", "Create edh-operational-alerts and confirm the email subscription.", "Subscription status is Confirmed."],
        ["7", "Create the six alarms with missing data set to not breaching. Add the alarm-status widget.", "Alarms are visible on EDH_Operational_Monitoring."],
        ["8", "Run the temporary exception test from the state machine, then remove the exception.", "ALARM email arrives, the error log contains the test text, and the next success returns the alarm to OK."],
        ["9", "Ask the ServiceNow owner whether the SES or operations mailbox already opens an incident.", "Option A or option B is chosen. The other path is not built."],
        ["10", "If option B is chosen, add the HTTPS subscription only after the URL is provided, and test one ALARM and one OK.", "One incident opens and the OK updates that same incident."],
    ],
)

h3("16.1 Permissions to request if a button is denied")
add_table(
    ["Action", "Permission", "Who grants it"],
    [
        ["Edit log retention", "logs:PutRetentionPolicy", "AWS administrator"],
        ["Create the Step Functions log delivery", "logs:CreateLogDelivery, logs:PutResourcePolicy, logs:DescribeLogGroups on the state machine role", "AWS administrator"],
        ["Create the SNS topic and subscriptions", "sns:CreateTopic, sns:Subscribe", "AWS administrator"],
        ["Create alarms and edit the dashboard", "cloudwatch:PutMetricAlarm, cloudwatch:PutDashboard", "AWS administrator"],
        ["Query logs", "logs:StartQuery, logs:GetQueryResults, logs:FilterLogEvents", "AWS administrator"],
        ["ServiceNow HTTPS endpoint", "The endpoint URL and the right to confirm the subscription", "ServiceNow administrator"],
    ],
)

note(
    "Keep the current foundation",
    "The 14 widgets, JobRunId ALL, the state machine ARN, and the bucket storage charts stay. "
    "Logs explain the run. Six alarms cover task failure, killed tasks, execution failure, timeout, "
    "finished-run delay, and heap pressure. SNS sends the alert. ServiceNow creates an incident only "
    "on the one path the ServiceNow owner already uses, or on the HTTPS subscription after that owner "
    "provides the endpoint.",
)

out = "/agent/repos/course/EDH_Dashboard_Review_Logs_Alarms_ServiceNow.docx"
doc.save(out)
print(out)
