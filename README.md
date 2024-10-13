# Send Calendar Email Script

## Overview
This VBA script automates sending calendar event emails from an Outlook account to a specified email address. It includes options to display detailed information for a given number of days.

## Script Breakdown
1. **Constants for Configuration:**
   Configurable options for the email address, inclusion of private details, and the number of days to display.
   ```vbscript
   Const myEmailAddress = "Email@email.com"
   Const includePrivateDetails = True
   Const howManyDaysToDisplay = 3
   ```

2. **Outlook Constants**: Pre-defined constants for calendar export formats and details.
   ```vbscript
   Const olCalendarMailFormatDailySchedule = 1
   Const olFreeBusyAndSubject = 1
   Const olFullDetails = 2
   Const olFolderCalendar = 9
   ```

3. **Send Calendar Function**: Main function that creates an Outlook application object, configures export settings, and sends the calendar event email.
   ```vbscript
   SendCalendar myEmailAddress, Date, (Date + (howManyDaysToDisplay - 1))

   Sub SendCalendar(strAdr, datBeg, datEnd)
       ' Function implementation...
   End Sub

   ```

## Usage

1. **Run the Script**: Execute the script in an Outlook VBA editor with appropriate permissions.


## License
This script is licensed under the MIT License. See the LICENSE file for details.


