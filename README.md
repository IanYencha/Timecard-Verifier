# Timecard-Verifier
Ian Yencha

## Description
This is a tool I made for to verfiy timecards as part of job I have at school. It compares the login times of each employee and checks it against the start and end of their shift to ensure they worked their scheduled hours.

### Input:
- **Schedule spreadsheet (.xlxs)**: needs employees username/fullname as the first row and column before their shifts. Each shift needs a date, start, and end time. 
- **Keys spreadsheet (.xlsx - could be another format with minor adjustments)**: needs username or full name, logon time (date + time) and logoff time (date + time)

### Output:
- **Formated Schedule spreadsheet (Verified.xlxs)**: Takes the schedule and highlights green for on time, yellow for late/left early, and red for no-show. If Yellow, the time they logon/logout will also be printed in col 7 and 8 respectivley. 

## Other Notes
This allows for a 5 minute buffer for being late/leaving early before being marked. It will also combine shifts that are close together. For example, consider an employee with a shift from 1p-3p and 3:30p-5p. Usually they do not logout for that 30 minutes and so I have choosen to just check for login at 1p and logoff at 5p. Any shifts within 1 hour of eachother will be combined this way. 

Overall, this program is a fairly niche tool, but if you find a need for it you're welcome to use it!
