# Process Tutoring
 
This repository contains a Google Cloud Function that processes and stores Calendly appointments to a Google Spreadsheet.   If an appointment is canceled or rescheduled, that information is reflected in the data.

This program was designed to track tutoring appointments in the College of Business at the University of Nebraska-Omaha. The expected columns are:

1. Student E-Mail	
2. Event Type	
3. Start Time	
4. End Time
5. Duration	
6. Canceled	
7. Calendly-UUID	
8. Class	
9. Instructor	
10. Topic
    
The last three columns come from questions added to the Calendly event questioner.  If these questions are not added, these fields will be left blank.

The program is driven by a service account file (service-account.json) -- you can obtain a service account JSON file by following the instructions [here](https://cloud.google.com/iam/docs/creating-managing-service-account-keys).  You need to add this service account as an editor to the target spreadsheet. 

## The Setting File

The settings.json file indicates the Calendly token, the default target spreadsheet, the data range, cancel range, id range, and if there is a custom spreadsheet for some users.  The file is relatively self explanatory.

## Creating a Google Cloud Function

You can create a Google Cloud Function by following the steps outlined in this [guide](https://cloud.google.com/functions/docs/quickstart-console): uploading all necessary code as a zip file.  This includes the service account JSON file. 

## Setting Up the Webhook

To setup the webhook, you need to tell Calendly where to send events.  You can do this by executing a single line from the terminal.

```bash
curl --request POST \
  --url https://api.calendly.com/webhook_subscriptions \
  --header 'Authorization: Bearer <Calendly Token>' \
  --header 'Content-Type: application/json' \
  --data '{
  "url": "https://<Cloud Function URL>",
  "events": [
    "invitee.created",
    "invitee.canceled"
  ],
  "organization": "<organization URI>",
  "user": "<user URI>",
  "scope": "organization"
}'
```

Where the token is your Calendly token and the cloud function URL is the URL provided by your new Cloud function.  The organization and user URIs are Calendly's unique identifiers for your user account or organization.  You can set the scope to "user" or "organization" depending on if you want to collect tutoring information just for yourself or all users in your organization.  Your Calendly account must be an admin user to use the organization scope.

## Getting Help

If you encounter any issues setting up this program, please contact Ben Smith at bosmith@unomaha.edu.