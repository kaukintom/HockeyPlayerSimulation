Option Explicit

Const year1 As Integer = 2017 'Change this to the most recently COMPLETED season
'Refers to year1/(year1 + 1) Season

'*************************************************************************************************
'***IMPORTANT***After this macro runs, the columns need to be rearranged so that they match previous seasons, _
'and the titles of each column need to match previous season titles
'*** Previous data was web scraped from a different source, so the way the tables are constructed are _
'different (Yahoo Sports vs. Sportsnet)
'*************************************************************************************************

Sub NHLStats() '****Web scrapes latest season of hockey player data
'This program may need updating due to a change in source code contents _
'and/or a change in web page configuration
'https://stackoverflow.com/questions/34703533/how-to-scrape-data-from-the-following-table-format-vba
'https://www.ozgrid.com/forum/forum/other-software-applications/excel-and-web-browsers-help/145319-select-an-item-from-a-dropdown-list-on-webpage
Dim IE As Object 'Internet explorer
Dim TableRow As Long 'Row
Dim TableCol As Long 'Column
Dim t As Long 'Child element
Dim SheetRow As Long 'Keeps track of what is the current row in the worksheet _
'for placing hockey player data
Dim RowStart As Long 'start place for scraping the data from the elements in HockeyStatsTable
Dim PagEnd As Boolean 'Pagination either available or not
Dim HockeyStatsTable As Object 'Hockey player data
Dim PagButton As Object 'Pagination
Dim NavPagRight As Object 'Pagination
Dim HockeyStatsDiv As Object 'Hockey player data
Dim DOMevent As Object 'Object to allow event to execute from drop down menu
Dim SeasonSelect As Object
Dim SheetToFind As String 'Worksheet for most recent hockey season
Dim sheet As Variant 'Contains every worksheet name

'Add  new worksheet and name it according to season
SheetToFind = year1 & "|" & (year1 + 1)

'Loops through worksheets to delete the most recent one if it already exists
For Each sheet In Worksheets
    If SheetToFind = sheet.Name Then
        Sheets(SheetToFind).delete
        Exit For
    End If
Next sheet

'Add the new sheet for the new season of hockey player data
Sheets.Add
ActiveSheet.Name = SheetToFind

Set IE = CreateObject("InternetExplorer.Application") 'Open internet explorer

With IE
    .Visible = True 'Show internet explorer
    .navigate ("https://www.sportsnet.ca/hockey/nhl/players/") 'Search URL _
    'for hockey player data
  
    'To make sure the webpage opens before executing the rest of the code _
    'put an loop that will be circular until IE has loaded the webpage
    Do While IE.readyState = 4: DoEvents: Loop
    Do Until IE.readyState = 4: DoEvents: Loop

    SheetRow = 0 'Starting value for row location on worksheet
    
    'Fires the onchange event
    'Selects a specific option from a dropdown menu
    Set DOMevent = IE.document.createEvent("HTMLEvents")
    DOMevent.initEvent "change", True, False
    
    Set SeasonSelect = IE.document.getElementById("season-dropdown") 'Dropdown menu name by ID
    SeasonSelect.Focus
    SeasonSelect.selectedIndex = 1  '2nd option from drop down menu (0 would be 1st option)
    SeasonSelect.dispatchEvent DOMevent 'Execute the event
    
    'Short time delay to allow the event to execute and webpage to load
    Application.Wait (Now + TimeValue("0:00:03"))
    
    Do 'Will loop until pagination limit reached
        
        'Find the element containing the Table of data for hockey players
        Do While HockeyStatsDiv Is Nothing
             Set HockeyStatsDiv = IE.document.getElementsByClassName("dataTables_scrollBody")
             DoEvents
        Loop
        
        Do While HockeyStatsDiv(0) Is Nothing 'Ensure that the elements exist within variable
             DoEvents
        Loop
        
        'Find where pagination occurs and with the element name
        Set PagButton = IE.document.getElementsByClassName("paginate_button next")
        'Child element for pagination navigating to the right
        Set NavPagRight = PagButton(0)
        
        'Name of element for when pagination is no longer available
        'If this is the name of the child element then PagEnd becomes true, signalling _
        'that there is no more data to be collected
        If NavPagRight.className = "paginate_button next disabled" Then PagEnd = True
        
        'Decides what data will be collected
        'When RowStart is = -1 then it will collect the row of data containing the headings of the columns
        If SheetRow = 0 Then RowStart = -1 Else RowStart = 0
        
        'Collects elements from table of hockey player data from the table within HockeyStatsDiv
        Set HockeyStatsTable = HockeyStatsDiv(0).getElementsByTagName("TABLE")
   
        'This portion finds each piece of data from the table and places it in a cell within _
        'the worksheet
        For TableRow = RowStart + 1 To (HockeyStatsTable(0).rows.Length - 1) 'Defines beginning position and _
        'ending position regarding the rows within table
            'Defines beginning position and ending position regarding the columns within table
            For TableCol = 0 To (HockeyStatsTable(0).rows(TableRow).Cells.Length - 1)
                'Placement of data within worksheet
                ThisWorkbook.ActiveSheet.Cells(TableRow + SheetRow + 1, TableCol + 1) = HockeyStatsTable(t).rows(TableRow).Cells(TableCol).innerText
            Next TableCol
        Next TableRow
    
        SheetRow = SheetRow + TableRow 'Saves row location within worksheet
    
        If Not NavPagRight Is Nothing Then NavPagRight.Click 'Execute pagination

        Set HockeyStatsDiv = Nothing 'Erase most recent hockey player data from variable

        Application.Wait (Now + TimeValue("0:00:01"))
        
    Loop Until PagEnd Or NavPagRight Is Nothing 'If pagination not avaiable then exit loop

End With

IE.Quit
Set IE = Nothing 'Close internet explorer

Call DeleteSpaces 'Deletes empty rows created from web scraping through pagination
Call FixNames   'Reformats names to fit existing format
End Sub

Sub DeleteSpaces()
'Deletes empty rows created from web scraping through pagination
Dim SeasonSheet As Worksheet 'Worksheet containing latest season of hockey player data
Dim SSarray As Variant 'Array containing contents of SeasonSheet
Dim SSrow As Integer, SScol As Integer 'Row and column location of SSarray and SeasonSheet

'Define which worksheet is SeasonSheet, put contents into array, and select the worksheet
SeasonSheet = Sheets(year1 & "|" & (year1 + 1))
SSarray = SeasonSheet.UsedRange
SeasonSheet.Select

'Loop through every row in SSarray
'Detects an empty cell, which means an empty row, and the row will be deleted from worksheet
For SSrow = LBound(SSarray) To UBound(SSarray, 1) 'Defines beginning and ending rows
    If IsEmpty(SSarray(SSrow, 1)) = True Then 'If the column of data containing _
    'the names of the hockey players has an empty entry
        Cells(SSrow, 1).EntireRow.delete 'Entire row is deleted
        SSarray = ActiveSheet.UsedRange 'Redefine array to update with deleted row
        SSrow = LBound(SSarray) 'Reset row location to the first entry
    End If
    If SSrow > UBound(SSarray, 1) Then Exit For
Next SSrow

End Sub

Sub FixNames()
Dim SeasonSheet As Worksheet 'Worksheet containing latest season of hockey player data
Dim SSarray As Variant 'Array containing contents of SeasonSheet
Dim SSrow As Integer, SScol As Integer 'Row and column variables for SeasonSheet cell location _
'and SSarray element location
Dim NameSplit As Variant 'Array of hockey player's name
Dim Name As String 'Hockey player's name

'Define which worksheet is SeasonSheet, put contents into array, and select the worksheet
SeasonSheet = Sheets(year1 & "|" & (year1 + 1))
SSarray = SeasonSheet.UsedRange
SeasonSheet.Select

'Loop through first column of SSarray
'First column contains names of hockey player's
For SSrow = LBound(SSarray) + 1 To UBound(SSarray, 1)
    NameSplit = Split(SSarray(SSrow, 1), ", ") 'Puts last and first names into an array
    Name = " " & NameSplit(1) & " " & NameSplit(0) 'Place first name before last name, _
    'with a space in the middle
    SSarray(SSrow, 1) = Name 'Replace name within array
Next SSrow

'Put array back onto worksheet with corrected name format
SeasonSheet.Cells(LBound(SSarray), LBound(SSarray)).Resize(UBound(SSarray, 1), UBound(SSarray, 2)) = SSarray
End Sub


