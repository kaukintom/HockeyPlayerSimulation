Option Explicit

Const year1 As Integer = 2017 'Change this to the most recent season
'Refers to year1-(year1 + 1) Season

'*************************************************************************************************
'***IMPORTANT***Before this macro runs, the columns need to be rearranged so that they match previous seasons, _
'and the titles of each column need to match previous season titles
'*************************************************************************************************


Sub Age_And_Season() '*****Web Scraper For Hockey Player Age
'Adds the age of the player and the current season of their career to the _
'Latest web scraped data
'If this is a player's first season, then their birthday will be scraped _
'from wikipedia
Dim SeasonSheet As Worksheet 'Latest season data
Dim SSarray As Variant 'Array of data on SeasonSheet
Dim Name As String 'Hockey Player
Dim Season As Integer 'Year that a hockey season began
Dim PrevSeason As Variant 'Array of data from a previous season
Dim SSrow As Integer 'Refers to a row for SSarray
Dim Prow As Integer 'Refers to a row for PrevSeason
Dim S As Variant 'Array for the season which might be split by a space
Dim Avalue As Integer 'Age of hockey player in previous season
Dim Svalue As Integer 'Number of seasons played before most recent season
Dim Aloc As Integer 'Location of Age column on worksheet
Dim Sloc As Integer 'Location of Season column on worksheet
Dim Age As Integer 'Age of selected hockey player
Dim CurrentYear As Integer 'Equal to year1
Dim SecondsElapsed As Double, StartTime As Double 'To determine length of program
Dim IE As Object 'Opens Internet Explorer

Application.ScreenUpdating = False

StartTime = Timer

CurrentYear = year1

Set SeasonSheet = Sheets(year1 & "|" & (year1 + 1))
SeasonSheet.Select

Set IE = CreateObject("InternetExplorer.Application") 'Open internet explorer

SSarray = SeasonSheet.UsedRange
'Create 2 new columns for the season and age data
Cells(1, UBound(SSarray, 2) + 1).Value = "Age"
Aloc = UBound(SSarray, 2) + 1
Cells(1, UBound(SSarray, 2) + 2).Value = "Season"
Sloc = UBound(SSarray, 2) + 2
'Loop through the names in the newest season to adjust their age and season _
', where the ages are adjusted to December 31
SSarray = SeasonSheet.UsedRange 'Array of latest season data

For SSrow = LBound(SSarray) + 1 To UBound(SSarray, 1) 'Add one to starting position to _
'account for the title headings of table
    Name = SSarray(SSrow, 1) ' Name of selected hockey player
    'Check the prior 10 season to see if player has played in the NHL, this is to adjust _
    'the season and age
    For Season = (year1 - 1) To (year1 - 10) Step -1
        Svalue = 0 'Reset value
        Avalue = 0 'Reset value
        'Make array of selected season to scan the contents quicker
        PrevSeason = Sheets(Season & "|" & (Season + 1)).UsedRange
        'Go through each row of selected season
        For Prow = LBound(PrevSeason) + 1 To UBound(PrevSeason, 1)
            'If the latest season they played is found, then obtain the value for their season
            If PrevSeason(Prow, 1) = Name Then
                'Create array for  season number since it might contain an "N"
                S = Split(PrevSeason(Prow, UBound(PrevSeason, 2)), " ")
                Svalue = S(LBound(S)) 'Number is contained in first element, "N" would be contained _
                'in the second element
                Avalue = PrevSeason(Prow, UBound(PrevSeason, 2) - 1) 'Age of player is one column to _
                'the left of the season value
                Exit For
            End If
        Next Prow
        'If their latest season is found, there is no need to look through past seasons
        If Avalue <> 0 And Svalue <> 0 Then Exit For
    Next Season
    'Change age value by the difference of the last time they played in the NHL
    SeasonSheet.Cells(SSrow, Aloc).Value = Avalue + (year1 - Season)
    'If their age value is zero then webscrape for age value
    If Avalue = 0 Then
        'Reset Age to zero
        Age = 0
        'Place on worksheet to indicate no age value
        SeasonSheet.Cells(SSrow, Aloc).Value = Avalue
        'Call web scraping code to get age value from wikipedia
        Call WikiBirthScraperV2(Name, Age, CurrentYear, IE)
        'Place discovered age value onto worksheet
        SeasonSheet.Cells(SSrow, Aloc) = Age
    End If

    'If they didnt play more than 25 games then the season number is accompanied by an N
    '25 is somewhat arbitrary, indicates not many games played in a season to have _
    'meaningful results evaluated from their stats in that season
    '25 is related to rookie player's ability to be viable for calder considerations
    If SSarray(SSrow, 3) >= 25 Then
        SeasonSheet.Cells(SSrow, Sloc).Value = Svalue + 1
    Else
        SeasonSheet.Cells(SSrow, Sloc).Value = (Svalue + 1) & " N"
    End If

Next SSrow

IE.Quit
Set IE = Nothing 'Close internet explorer

Application.ScreenUpdating = True

MsgBox (Round(Timer - StartTime, 2) / 60 & " Minutes.")

End Sub

Sub WikiBirthScraperV2(Name, Age, CurrentYear, IE)
'Extract birthday data from wikipedia
Dim NameSplit As Variant 'When name is split, the parts of the name are in this array
Dim NAcol As Integer 'Refers to element number in NameSplit
Dim URLname As String 'Constructed portion within URL
Dim URL As String 'Web page to open within Internet Explorer
Dim BdayObj As HTMLDivElement 'Refers to html source code elements on web page
Dim Birthday As String 'yyyy-mm-dd of hockey player from web page source code
Dim NAMESarray As Variant 'Array of contents from "NAMES" worksheet
Dim Nsheet As Worksheet 'NAMES worksheet
Dim BirthdayCol As String 'Name of column heading containing birthday data
Dim NameCol As String 'Name of column heading containing Name data
Dim URLcol As String 'Name of column heading containing URL data
Dim Nrow As Integer, Ncol As Integer 'Element location for NAMESarray _
'refering to row and column number in array

Application.Wait (Now + TimeValue("00:00:01")) 'Time delay to prevent _
'any problems from accessing the site too quickly

NameSplit = Split(Name, " ") 'Put the parts of the name into one dimensional array

Set Nsheet = Sheets("NAMES")

NAMESarray = Nsheet.UsedRange 'List of hockey player's, birthdays, and wikipedia URL's

'Column names on Nsheet
NameCol = "Name"
URLcol = "URL"
BirthdayCol = "Birthday"

'Put together a string that is part of the URL
'Wikipedia URL follows the format: https://en.wikipedia.org/wiki/FirstName_LastName
'This for loop creates the FirstName_LastName portion
For NAcol = LBound(NameSplit) To UBound(NameSplit) 'Loop through each part of name
    If NAcol = UBound(NameSplit) Then
        URLname = URLname & NameSplit(NAcol) 'Add name --> Dont want an underscore since _
        'this is the player's last name
    Else 'Add Name with an underscore at the end
        URLname = URLname & NameSplit(NAcol) & "_" 'Not player's lsat name, or last part _
        'of last name
    End If
Next NAcol

URL = "https://en.wikipedia.org/wiki/" & URLname 'wikipedia URL for hockey player

'Since this object was created in a previous sub procedure, the object can be refered to
With IE
    .Visible = False 'Do not show internet explorer
    'If this sub is terminated before reaching the end, an invisibale instance of _
    'Internet Explorer will still exist, using computing resources
    .navigate URL 'Search URL for hockey player
  
    'To make sure the webpage opens before executing the rest of the code _
    'put an loop that will be circular until IE has loaded the webpage
    Do While IE.readyState = 4: DoEvents: Loop
    Do Until IE.readyState = 4: DoEvents: Loop
    
    'These are fixed elements that exist for every hockey player's page on wikipedia
    'Thay can be accessed from the wikipedia page using inspect on the webpage
    'There have been 3 exceptions where this code has encountered an error _
    'due to an hockey player's wikipedia page no longer existing
    Set BdayObj = IE.document.getElementById("mw-content-text")
    Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
    
    'If the element "Infobox vcard" does not exist then an additional sting needs _
    'to be added to the URL and searched
    If BdayObj Is Nothing Then ' "Infobox vcard" does not exist --> BdayObj is nothing
        URL = "https://en.wikipedia.org/wiki/" & URLname & "_(ice_hockey)" ' Added "_(ice_hockey)"
        'There are multiple people with wikipedia profiles with the same name
        .navigate URL 'Search URL for hockey player data
        
        
        'To make sure the webpage opens before executing the rest of the code _
        'put an loop that will be circular until IE has loaded the webpage
        Do While IE.readyState = 4: DoEvents: Loop
        Do Until IE.readyState = 4: DoEvents: Loop
            
        'Explained above as above
        Set BdayObj = IE.document.getElementById("mw-content-text")
        Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
    End If
    
    If Not BdayObj Is Nothing Then 'If 'Inforbox vcard' now exists
        Set BdayObj = BdayObj.getElementsByTagName("td")(1) 'This takes the second child element _
        'with the tagname "td"
        If Not BdayObj Is Nothing Then 'If there is a second "td"
            Set BdayObj = BdayObj.getElementsByClassName("bday")(0) 'Then want to set BdayObj to the _
            'element containing the hockey player's birthday where the class name is "bday"
            If Not BdayObj Is Nothing Then 'If "bday exists" get the birthday data within the _
            'html data
                Birthday = BdayObj.innerHTML
            Else 'If there was no element named "bday"
                'Entire process is same except for line 3
                Set BdayObj = IE.document.getElementById("mw-content-text")
                Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
                Set BdayObj = BdayObj.getElementsByTagName("td")(0) 'This line is changed to refer to the _
                'first child element "td" and not the second
                Set BdayObj = BdayObj.getElementsByClassName("bday")(0)
                If Not BdayObj Is Nothing Then
                    Birthday = BdayObj.innerHTML
                End If
            End If
        End If
    End If
    
    'Add name to Nsheet
    Nsheet.Cells(UBound(NAMESarray, 1) + 1, 1).Value = Name
    'Redefine NAMESarray to include newly added player
    NAMESarray = Nsheet.UsedRange
    
    'If a birthday was extracted from wikipedia
    If IsEmpty(Birthday) = False And Birthday <> "" Then 'If for some reason there is no birthday value _
    'yet the variable is not being retrieved as empty
        'Loop through column number in array
        For Ncol = LBound(NAMESarray) To UBound(NAMESarray, 2)
            If NAMESarray(1, Ncol) = URLcol Then 'If the column containing URLs is found
                Nsheet.Cells(UBound(NAMESarray, 1), Ncol).Value = URL 'Place URL in worksheet
            ElseIf NAMESarray(1, Ncol) = BirthdayCol Then 'If the column containing birthdays is found
                Nsheet.Cells(UBound(NAMESarray, 1), Ncol).Value = Birthday 'Place birthday in worksheet
                Exit For 'Since Birthday is found after URL
            End If
        Next Ncol
    End If
End With

'If a birthday was retrieved, find the age of the player as of December 31 during the season
If IsEmpty(Birthday) = False And Birthday <> "" Then
    Call AdjustBirthday(Birthday, Age, CurrentYear)
End If

End Sub

Sub AdjustBirthday(Birthday, Age, CurrentYear)

Dim BdayArray As Variant

BdayArray = Split(Birthday, "-") 'Put birthday string into array

Age = CurrentYear - BdayArray(0) 'Current year subtracted by the first element in array (birth year)

End Sub

'Sub WikiBirthScraperV2_Variant()
'*****This was used to go through the list of extracted URL's to use the most current webscraping _
'method to extract all the birthdays of the hockey players
'****This was a testing piece of code
'Dim NAMEarray As Variant
'Dim NAcol As Integer
'Dim URLname As String
'Dim URL As String
'Dim BdayObj As HTMLDivElement
'Dim Birthday As String
'Dim NAMESarray As Variant
'Dim Nsheet As Worksheet
'Dim BirthdayCol As String
'Dim NameCol As String
'Dim URLcol As String
'Dim Nrow As Integer, Ncol As Integer
'Dim Name As String
'Dim SecondsElapsed As Double, StartTime As Double
'Dim IE As Object, ColNum As Integer
'Dim i As Integer, j As Integer, pctCompl As Single
'Dim Count As Integer, Total As Integer
'
'Application.ScreenUpdating = False
'
'StartTime = Timer
'
'Count = 0
'
'Set Nsheet = Sheets("NAMES")
'
'NAMESarray = Nsheet.UsedRange
'Nsheet.Select
'
'NameCol = "Name"
'URLcol = "URL"
'BirthdayCol = "Birthday"
'
'For Ncol = LBound(NAMESarray) To UBound(NAMESarray, 2)
'    If NAMESarray(1, Ncol) = URLcol Then Exit For
'Next Ncol
'ColNum = Ncol
'
'Set IE = CreateObject("InternetExplorer.Application") 'Open internet explorer
'
'For Nrow = LBound(NAMESarray) + 1 To UBound(NAMESarray, 1)
'
'    URLname = ""
'    Name = NAMESarray(Nrow, 1)
'    NAMEarray = Split(Name, " ")
'    Birthday = ""
'    Set BdayObj = Nothing
'
'    For NAcol = LBound(NAMEarray) To UBound(NAMEarray) 'Loop through each part of name
'        If NAcol = UBound(NAMEarray) Then
'            URLname = URLname & NAMEarray(NAcol) 'Add name --> Dont want an underscore since _
'            this is the player's last name
'        Else 'Add Name with an underscore at the end
'            URLname = URLname & NAMEarray(NAcol) & "_" 'Not player's lsat name, or last part _
'            of last name
'        End If
'    Next NAcol
'
'    URL = "https://en.wikipedia.org/wiki/" & URLname 'wikipedia URL for hockey player
'
'    With IE
'        .Visible = False 'Show internet explorer
'        .navigate URL 'Search URL for hockey player data
'
'        'To make sure the webpage opens before executing the rest of the code _
'        put an loop that will be circular until IE has loaded the webpage
'        Do While IE.readyState = 4: DoEvents: Loop
'        Do Until IE.readyState = 4: DoEvents: Loop
'        'Eric Nickulas
'        'Peter Sarno
'        'Jason Doig
'        '------------
'        'Harold Druken
'        'Dan Kesa
'        'Stop
'        If (Name <> " Greg Moore") Or (Name <> " Michael Ryan") Or (Name <> " Craig MacDonald") Then
'            On Error Resume Next
'            Set BdayObj = IE.document.getElementById("mw-content-text")
'
'            If Not BdayObj Is Nothing Then
'                Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
'                'Stop
'                'Stop
'                If Not BdayObj Is Nothing Then
'                    Set BdayObj = BdayObj.getElementsByTagName("td")(1)
'                    'Set BdayObj = Nothing
'                    'Stop
'                    If Not BdayObj Is Nothing Then
'                        Set BdayObj = BdayObj.getElementsByClassName("bday")(0)
'                        If Not BdayObj Is Nothing Then
'                            Birthday = BdayObj.innerHTML
'                        Else
'                            Set BdayObj = IE.document.getElementById("mw-content-text")
'                            Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
'                            Set BdayObj = BdayObj.getElementsByTagName("td")(0)
'                            Set BdayObj = BdayObj.getElementsByClassName("bday")(0)
'                            If Not BdayObj Is Nothing Then
'                                Birthday = BdayObj.innerHTML
'                            End If
'                            'Stop
'                        End If
'                    End If
'                End If
'            End If
'        End If
'
'        If BdayObj Is Nothing Then
'            'Stop
'            URL = "https://en.wikipedia.org/wiki/" & URLname & "_(ice_hockey)" ' Added "_(ice_hockey)"
'
'            .navigate URL 'Search URL for hockey player data
'                    'To make sure the webpage opens before executing the rest of the code _
'        put an loop that will be circular until IE has loaded the webpage
'            Do While IE.readyState = 4: DoEvents: Loop
'            Do Until IE.readyState = 4: DoEvents: Loop
'
'            'Explained above
'            If (Name <> " Greg Moore") Or (Name <> " Michael Ryan") Or (Name <> " Craig MacDonald") Then
'                On Error Resume Next
'                Set BdayObj = IE.document.getElementById("mw-content-text")
'                'Stop
'                If Not BdayObj Is Nothing Then
'                    Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
'                    'Stop
'                    'Stop
'                    If Not BdayObj Is Nothing Then
'                        Set BdayObj = BdayObj.getElementsByTagName("td")(1)
'                        'Set BdayObj = Nothing
'                        'Stop
'                        If Not BdayObj Is Nothing Then
'                            Set BdayObj = BdayObj.getElementsByClassName("bday")(0)
'                            'Stop
'                            If Not BdayObj Is Nothing Then
'                                Birthday = BdayObj.innerHTML
'                                'Stop
'                            Else
'                                Set BdayObj = IE.document.getElementById("mw-content-text")
'                                Set BdayObj = BdayObj.getElementsByClassName("infobox vcard").Item(0)
'                                Set BdayObj = BdayObj.getElementsByTagName("td")(0)
'                                Set BdayObj = BdayObj.getElementsByClassName("bday")(0)
'                                If Not BdayObj Is Nothing Then
'                                    Birthday = BdayObj.innerHTML
'                                End If
'                                'Stop
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'
'        If IsEmpty(Birthday) = False And Birthday <> "" Then
'            Count = Count + 1
'            For Ncol = LBound(NAMESarray) To UBound(NAMESarray, 2)
'                If NAMESarray(1, Ncol) = URLcol Then
'                    If Birthday <> "" Then Nsheet.Cells(Nrow, Ncol).Value = URL
'                ElseIf NAMESarray(1, Ncol) = BirthdayCol Then
'                    If Birthday <> "" Then Nsheet.Cells(Nrow, Ncol).Value = Birthday
'                    Exit For 'Since Birthday is found after URL
'                End If
'            Next Ncol
'        End If
'    End With
'
'    pctCompl = Round((Nrow - 1) / (UBound(NAMESarray, 1) - 1) * 100, 2)
'    progress pctCompl
'    Application.Wait (Now + TimeValue("00:00:01"))
'Next Nrow
'
'Set IE = Nothing 'Close internet explorer
'
'Application.ScreenUpdating = True
'
'Total = (UBound(NAMESarray, 1) - 1)
'MsgBox (Round(Timer - StartTime, 2) / 60 & " Minutes. Names Scraped = " & Count & " out of " & Total)
''Add one to result
'End Sub

'Sub RemainingAges()
'****Find age of players as of december 31 during the most recent season
'****Only for most recent Season
'****This was a testing piece of code
'Dim SeasonSheet As Worksheet
'Dim SSarray As Variant, Name As String
'Dim SSrow As Integer, SScol As Integer
'Dim NAMESarray As Variant
'Dim Nsheet As Worksheet
'Dim NameCol As Integer, BdayCol As Integer
'Dim Col As Integer, Birthday As String
'Dim DateArray As Variant
'Dim CurrentYear As Integer
'Dim Nrow As Integer
'
'CurrentYear = year1
'
'Set SeasonSheet = Sheets(year1 & "|" & (year1 + 1))
'Set Nsheet = Sheets("NAMES")
'
'SeasonSheet.Select
'
'SSarray = SeasonSheet.UsedRange
'NAMESarray = Nsheet.UsedRange
'
'For Col = LBound(NAMESarray) To UBound(NAMESarray, 2)
'    If NAMESarray(1, Col) = "Name" Then
'        NameCol = Col
'    ElseIf NAMESarray(1, Col) = "Birthday" Then
'        BdayCol = Col
'    End If
'Next Col
'
'For SScol = LBound(SSarray) To UBound(SSarray, 2)
'    If SSarray(1, SScol) = "Age" Then
'        For SSrow = LBound(SSarray) + 1 To UBound(SSarray, 1)
'            Name = SSarray(SSrow, NameCol)
'            For Nrow = LBound(NAMESarray) To UBound(NAMESarray, 1)
'                If NAMESarray(Nrow, NameCol) = Name Then
'                    Birthday = NAMESarray(Nrow, BdayCol)
'                    DateArray = Split(Birthday, "-")
'                    SeasonSheet.Cells(SSrow, SScol).Value = CurrentYear - DateArray(0)
'                    Exit For
'                End If
'            Next Nrow
'        Next SSrow
'    End If
'Next SScol
'End Sub


Sub progress(pctCompl As Single)

Userform1.Text.Caption = pctCompl & "% Completed"
Userform1.Bar.Width = pctCompl * 2

DoEvents

End Sub

