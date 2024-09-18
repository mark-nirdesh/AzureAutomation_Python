
# Azure Function for Automated OneDrive File Processing and Notification

This project automates the process of downloading an Excel timesheet file from OneDrive, modifying its contents, and sending it via email using Azure Functions. Additionally, it sends a notification to a Matrix room to confirm the completion of the task. The function is triggered every Friday at 4:30 PM using Azure's Timer Trigger, making it a reliable and scheduled automation solution.

## Features

- **Automated File Download**: Uses Microsoft Graph API to download an Excel file from a specific OneDrive location.
- **Excel File Processing**: Modifies specific cells in the Excel file to reflect the previous Monday and next Sunday based on the current date.
- **Email Sending**: Automatically sends the modified Excel file as an attachment via SMTP using Outlook.
- **Matrix Notification**: Sends a notification to a specified Matrix room to confirm the successful execution of the task.
- **Scheduled Execution**: Runs every Friday at 4:30 PM using Azure Functions Timer Trigger, ensuring timely updates without manual intervention.
- **Error Handling and Logging**: Includes detailed logging and error handling for easier debugging and monitoring.

## Technologies Used

- **Azure Functions**: Serverless compute service for executing the automation on a schedule.
- **Python**: The primary programming language used for the function.
- **Microsoft Graph API**: For accessing and downloading files from OneDrive.
- **OpenPyXL**: For reading and modifying Excel files.
- **SMTP (smtplib)**: For sending emails through Outlook.
- **Matrix API**: For sending notifications to a Matrix room.
- **MSAL (Microsoft Authentication Library)**: For authenticating and acquiring access tokens to access OneDrive.

## Project Structure

- **function_app.py**: Main Python script containing the Azure Function logic.
- **config.json**: Configuration file storing sensitive information like email credentials and Matrix API tokens (ensure this is secured and not exposed publicly).
- **requirements.txt**: List of dependencies required to run the function on Azure.

## Setup and Deployment

### Prerequisites

- **Azure Subscription**: To create and deploy Azure Functions.
- **OneDrive Personal Account**: For storing the Excel file.
- **SMTP Account (Outlook)**: For sending emails.
- **Matrix Account**: For sending notifications to a Matrix room.

### How to Deploy

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/mark-nirdesh/AzureAutomation_Python.git
   ```
   
2. **Configure Environment**:
   - Update `config.json` with the necessary credentials and settings.
   - Ensure that sensitive data such as client IDs, secrets, and tokens are securely managed.

3. **Deploy to Azure**:
   - Use Azure Functions Core Tools or Visual Studio Code to deploy the function to Azure.
   - Ensure that the function app is set to use Python runtime.
   - Configure CORS settings in the Azure Portal to allow `https://portal.azure.com` if you want to trigger tests directly from the portal.

4. **Monitor and Test**:
   - Use Azure Portal to monitor the function's execution.
   - Manually trigger the function from the portal for testing purposes.

## Usage

- The function runs automatically every Friday at 4:30 PM.
- On each execution:
  - Downloads the specified Excel timesheet from OneDrive.
  - Modifies the file to update the dates in specific cells.
  - Sends an email with the modified Excel file attached.
  - Sends a notification to a Matrix room.

## Future Enhancements

- **Error Notification**: Add functionality to notify administrators in case of errors during execution.
- **Enhanced Security**: Implement Azure Key Vault to securely manage and access credentials.
- **Dynamic Scheduling**: Allow dynamic changes to the schedule via an HTTP trigger or configuration updates.
- **Logging and Monitoring**: Integrate with Application Insights for more detailed logging and telemetry data.


## Contributions

Contributions are welcome! Please open an issue or submit a pull request with your improvements.
