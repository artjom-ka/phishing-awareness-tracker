# Phishing Awareness Tracker

An automated serverless system that monitors employee security awareness 
training progress via the Wizer API and automatically notifies HR about 
employees who haven't completed their assigned courses.

## How it works

1. **EventBridge Scheduler** triggers the Lambda function on a defined 
   schedule (e.g. twice a week)
2. **Lambda** retrieves the Wizer API key securely from **Secrets Manager**
3. Employee training data is fetched from the Wizer Master Report API
4. Each employee is categorized into one of three groups:
   - `diligent` — all courses completed (progress = 100%)
   - `undiligent` — no courses started or all statuses incomplete
   - `mixed` — partially completed
5. Results are saved to an **Excel file** in **S3**, with a new sheet 
   created on every run (up to 5 sheets retained)
6. Changes are tracked between runs — new undiligent users and users 
   who improved their status are detected automatically
7. A structured report is published to **SNS**, which emails it to 
   subscribed HR staff automatically

## Requirements

The following libraries are required. `boto3` comes pre-installed 
in the Lambda runtime — the rest must be packaged with the deployment ZIP.

```
boto3        # AWS SDK — pre-installed in Lambda
requests     # HTTP requests to Wizer API
openpyxl     # Excel file creation and editing
croniter     # Cron expression parsing for next-run calculation
```

## Deployment

1. Install dependencies into a local directory:
```bash
pip install requests openpyxl croniter -t ./package
cp lambda_function.py ./package/
```

2. Create a ZIP archive:
```bash
cd package
zip -r ../phishing-awareness-tracker.zip .
```

3. Upload the ZIP to your AWS Lambda function via the AWS Console 
   or AWS CLI

4. Set the following values inside `lambda_handler` to match 
   your environment:

| Variable | Description |
|---|---|
| `region_name` | Your AWS region (e.g. `eu-central-1`) |
| `bucket_name` | S3 bucket name for Excel storage |
| `secret_name` | Secrets Manager secret name for Wizer API key |
| `topic_arn` | SNS topic ARN for HR notifications |
| `schedule_name` | Your EventBridge Scheduler name |
