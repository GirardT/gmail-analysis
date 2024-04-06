# Gmail Analysis Tool

This tool is designed to analyze emails in your Gmail inbox and provide insights into the distribution of emails by sender domain. It fetches emails from your inbox, counts the number of emails received from each sender domain, and presents the analysis in a Google Sheets document.

## Features

- Fetches emails from the inbox using the Gmail API.
- Handles sender names with special characters like "@" in their email addresses.
- Provides options to retrieve search patterns or list emails with counts.
- Presents the analysis in a Google Sheets document.
- Offers a simple menu option in Google Sheets for easy execution.

## Usage

1. **Set Up Google Sheets:**
   - Create a new Google Sheets document or open an existing one.
   - Rename the first sheet to "Email Analysis".

2. **Set Up Gmail API:**
   - Enable the Gmail API for your Google Cloud project.
   - Ensure the necessary scopes and permissions are granted.

3. **Install the Script:**
   - Open the Google Sheets document.
   - From the toolbar, go to Extensions > Apps Script.
   - Replace the contents of the `Code.gs` file with the provided script.
   - Save the script.

4. **Run the Script:**
   - Close the Apps Script editor.
   - Refresh the Google Sheets document.
   - You should now see a new menu option named "Count Senders" under the "Add-ons" menu.
   - Click on "Count Senders" to expand the menu, and choose either "Get Search Patterns" or "List Emails with Count" to execute the script.

5. **View Results:**
   - The script will fetch emails from your inbox and analyze them.
   - If you chose "Get Search Patterns", the search patterns will be displayed in the "Email Analysis" sheet.
   - If you chose "List Emails with Count", the sender domains and their corresponding email counts will be displayed in the "Email Analysis" sheet.

## Notes

- This tool fetches a maximum of 1000 emails from your inbox. You can adjust this limit in the code if needed.
- Ensure that the necessary permissions are granted to access your Gmail inbox and Google Sheets document.
- You may need to enable advanced Google services and APIs in the Google Cloud Console.
- For any issues or feedback, please open an issue on the GitHub repository.
