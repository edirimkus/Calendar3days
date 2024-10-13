' -------------------------------------------------
' Modify this \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Const myEmailAddress = "Email@email.com"
Const includePrivateDetails = True
Const howManyDaysToDisplay = 3
' Modify this /////////////////////////////////////
' -------------------------------------------------
Const olCalendarMailFormatDailySchedule = 1
Const olFreeBusyAndSubject = 1
Const olFullDetails = 2
Const olFolderCalendar = 9
SendCalendar myEmailAddress, Date, (Date + (howManyDaysToDisplay - 1))
Sub SendCalendar(strAdr, datBeg, datEnd)
Dim olkApp, olkSes, olkCal, olkExp, olkMsg
Set olkApp = CreateObject("Outlook.Application")
Set olkSes = OlkApp.GetNameSpace("MAPI")
olkSes.Logon "Default"
Set olkCal = olkSes.GetDefaultFolder(olFolderCalendar)
Set olkExp = olkCal.GetCalendarExporter
With olkExp
.CalendarDetail = olFreeBusyAndSubject
.IncludePrivateDetails = includePrivateDetails
.RestrictToWorkingHours = False
.StartDate = datBeg
.EndDate = datEnd
End With
Set olkMsg = olkExp.ForwardAsICal(olCalendarMailFormatDailySchedule)
With olkMsg
.To = strAdr
.Send
End With
Set olkCal = Nothing
Set olkExp = Nothing
Set olkMsg = Nothing
olkSes.Logoff
Set olkSes = Nothing
Set olkApp = Nothing
End Sub
