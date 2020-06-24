The Excel file is a macro-enabled workbook with code.  The workbook contains a set of functions used to login into
the Jira REST API and query it for data.  The workbook makes use of VBA-JSON v2.3.1. 
(c) Tim Hall - https://github.com/VBA-tools/VBA-JSON

VBA-JSON is used to parse the JSON results returned from the Jira REST API into VB-specific data structures (e.g.
arrays and dictionaries) that are much easier for VB developers to work with than JSON strings.

For more information about VBA-JSON and the JsonCoverter module refer to the github location above.

To use  the macro:
Add the URL for your instance of the Jira REST API and Jira REST API Authorization on the configuration page.
The functions in the VB code will use these values to send authentication information to your Jira server instanace
and query the REST API for data on project(s) you have rights to view.

For more information on the Jira REST API and other things you can do with it you can visit
https://developer.atlassian.com/server/jira/platform/rest-apis/