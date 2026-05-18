# Payroll Query System

A web form for workers to submit payroll queries directly to the accounts team via email.

## How it works

```mermaid
flowchart TD
    W[Worker fills out form] --> F[Flask app - app.py]
    F --> V[Validate fields + attachments]
    V --> G[Microsoft Graph API]
    G --> E[Email sent to accounts team\naccounts@marlugroupwa.com.au]

    classDef blue fill:#dbeafe,stroke:#2563eb,color:#1e40af
    classDef green fill:#d1fae5,stroke:#059669,color:#064e3b
    classDef amber fill:#fef3c7,stroke:#d97706,color:#92400e
    classDef gray fill:#f3f4f6,stroke:#6b7280,color:#111827

    class W amber
    class F,V gray
    class G blue
    class E green
```

## What the form collects

| Field | Required |
|---|---|
| Full name | Yes |
| Phone number | Yes |
| Email address | Yes |
| Employer | Yes |
| Site | No |
| Pay period (start + end) | Yes |
| Query type | Yes |
| Description | No |
| Attachments (PDF, JPG, PNG, DOC, DOCX, XLS, XLSX) | No |

## Attachment Rules

- Max **3 files** per submission
- Max **20MB** total size
- Allowed types: PDF, JPG, PNG, DOC, DOCX, XLS, XLSX

## Technologies

- Python 3.11
- Flask
- Microsoft Graph API (send email via Office 365)
- Azure App Service (deployment)
- GitHub Actions (CI/CD)

```
## Project Structure
payroll-query-system/
├── app.py              Flask app + Graph API email logic
├── templates/
│   └── form.html       Payroll query web form
├── requirements.txt
├── startup.txt
└── .github/workflows/  Azure App Service deployment
```

## Environment Variables

| Variable | Description |
|---|---|
| GRAPH_TENANT_ID | Azure AD tenant ID |
| GRAPH_CLIENT_ID | Azure AD app client ID |
| GRAPH_CLIENT_SECRET | Azure AD app client secret |
| GRAPH_SENDER | Sender email address |
| GRAPH_TO | Recipient email (comma separated for multiple) |
