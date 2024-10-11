# D365Utility

Analyzing Field Usage in Dynamics 365 Forms and Views with C#
In this post, we explore a C# console application that retrieves the FetchXML of forms and views present in a Dynamics 365 environment to check if specific fields are used in them. This is useful for analyzing field usage and cleaning up unused attributes in your CRM system. Let’s break down the key parts of the code.

Key Highlights of the Code
Dynamics 365 Connection:
The application fetches necessary connection parameters such as Url, ClientId, and ClientSecret from the app.config file to authenticate with Dynamics 365 using the ServiceClient.

FetchXML for Forms and Views:
Two FetchXML queries are used:

Form Fetch: Retrieves form-related data, including form JSON and XML.
View Fetch: Retrieves saved queries (views) data, including fetch XML and related attributes.
Reading Field Data from Excel:
The application reads an Excel file containing field details (table name and logical name). This data is used to check if fields are utilized in any form or view within the system.

Field Usage Analysis:
For each field, the code checks both the form and view FetchXML to determine if the field is referenced. If found, the form or view's name and ID are recorded. Regular expressions sanitize the XML data for easy searching.

Result Logging and Progress Tracking:
Progress is logged as the application checks each field. It also logs when a certain number of fields are processed to keep track of the operation's pace.

Saving Results to Excel:
The result of the field usage analysis is saved in an Excel file, with columns indicating the table name, attribute name, whether the field is used, and where it was found (form or view).

Conclusion
This simple yet powerful application allows CRM developers and administrators to quickly determine whether fields are being utilized in forms or views, aiding in system maintenance and optimization.
