# PSServiceNow
PowerShell Functions for the ServiceNow API

![alt text](https://github.com/hematic/Storage/raw/master/servicenow.jpg)

# Requirements
 - Know the URL for your ServiceNow Instance
 - Valid Credentials for your ServiceNow instance.
 - PowerShell 3.0 or greater
# Features
  - Retrieval of a ServiceNow user
    - By Samaccountname 
    - By Email Address
    - By First and Last Name
  - Retrieval of ServiceNow Incident
    - By IncidentID
    - By Description
    - By ServiceNow SYSUserID
    - By Assigned to Email Address
    - By First and Last name of assigned user
  - Create a ServiceNow Incident
    - Passing of almost any field is available however its possible your organization has customized your incident form to use different field names. If so you may need to customize this function to use the fields your organization uses. To do so just simply change the function parameter names. The help example shows the fields available.
    - By Iteration Name
  - Update a ServiceNow Incident
    - Passing of almost any field is available however its possible your organization has customized your incident form to use different field names. If so you may need to customize this function to use the fields your organization uses. To do so just simply change the function parameter names. The help example shows the fields available.

# Examples and Help

Currently located inside the functions themselves.
