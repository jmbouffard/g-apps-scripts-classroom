# g-apps-scripts-classroom
Collection of Google Apps scripts used to automate tasks in a classroom environment

## Data2Doc
This Google Sheets script creates a document from the values in predefined columns. The lines are selected based on the date selected by the user. Everytime the script is ran for a new date, a new document is created. A template is used to create the new document and the script fails if the template is not found.

*The script was developed using a specific Sheets document and would have to be adapted to use with a different format.*

## EasySendEmail
This Google Sheets script sends an email with the content of each line. The email address used is located in column "emailAddress". An email template must be included in a 2nd sheet. Everytime the script is ran, emails will be sent for all lines of the document.

*The script was developed using a specific Sheets document and would have to be adapted to use with a different format.*

## DocOnSubmit
This Google Forms script appends a table with all answers to the form to a Docs document. The document name must be included as the answer to the first question of the form. If the document does not exist it will be created before being populated. The script runs everytime the form is submitted.

*The script was developed using a specific Forms document and would have to be adapted to use with a different format.*

