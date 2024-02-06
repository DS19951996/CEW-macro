

Option Explicit
Dim fileName As String
Dim PreviousweekDate
Dim DirectoryPath As String
Dim Path As String
Sub FormatInputFile()
'
' Macro2 Macro
'
   

    DirectoryPath = ThisWorkbook.Sheets("Macro").Range("B3")
    fileName = "AllSuspendedChargesForAnalysis_" & Format(Now, "yyyyMMDD") & ".xlsx"
    PreviousweekDate = Format(DateAdd("ww", -1, Format(Now, "yyyy-MM-DD")), "yyyy-MM-DD")
    Path = DirectoryPath & "AllSuspendedChargesForAnalysis_" & Format(Now, "yyyyMMDD") & ".xlsx"
    Workbooks.Open (Path)
    
    Workbooks(fileName).Activate
    
    Worksheets("AllSuspendedChargesForAnalysis1").Activate

    'Worksheets("AllSuspendedChargesForAnalysis1").Activate'
    'Windows("AllSuspendedChargesForAnalysis_" & Format(Now(), "YYYYMMDD") & ".xlsx").Activate'
    Columns("D:D").Select
    Selection.TextToColumns Destination:=Range("D1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    
    Columns("J:J").Select
    Selection.TextToColumns Destination:=Range("J1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Columns("S:S").Select
    Selection.TextToColumns Destination:=Range("S1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
        
        


End Sub
Sub CopyNewDatatoCurrentFile()
'
' Macro3 Macro
'

'
    
   
    fileName = "AllSuspendedChargesForAnalysis_" & Format(Now, "yyyyMMDD") & ".xlsx"
    
    Dim SuspenseErrorAnalysisFilePath
    Dim wb
    Dim lRowTC
    Dim lRow
    
    
    'Workbooks("C:\Users\divyajain24\Desktop\Suspense Reports\Macro.xlsm").Activate
    Path = ThisWorkbook.Sheets("Macro").Range("B1")
    SuspenseErrorAnalysisFilePath = DirectoryPath & PreviousweekDate & " Suspense Error Analysis.xlsx"
   Application.DisplayAlerts = False
    Set wb = Workbooks.Open(SuspenseErrorAnalysisFilePath)
    wb.SaveAs fileName:=DirectoryPath & Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx", FileFormat:=xlOpenXMLStrictWorkbook
    
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Trending Counts").Activate
    lRowTC = Cells(Rows.count, 1).End(xlUp).Row
    
    
    
   ' lRowTC = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    Range("A" & lRowTC + 1).Select
    Dim K As String
    K = Date
    K = Format(K, "MM/DD/yy")
    ActiveCell.FormulaR1C1 = K
    Range("A" & lRowTC + 1).Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        

    


    
    
    Worksheets("Detail").Activate
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    
    Dim lRowNew
     
    lRow = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    
    'lRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    
    
    ActiveSheet.Range("A3:AQ" & lRow).Select
    
    
    'Application.CutCopyMode = False
    Selection.ClearContents
    
    Workbooks(fileName).Activate
    
     
    lRowNew = Cells.Find(What:="*", _
                    After:=Range("A1"), _
                    LookAt:=xlPart, _
                    LookIn:=xlFormulas, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlPrevious, _
                    MatchCase:=False).Row
    
    Worksheets("AllSuspendedChargesForAnalysis1").Range("A1:AC" & lRowNew).Copy
    'ActiveWindow.LargeScroll ToRight:=-1
    'Range("A1").Select
    
    
    'Range(Selection, Selection.End(xlDown)).Select
    'Range(Selection, Selection.End(xlToRight)).Select
    'Selection.Copy
    'Windows("2021-11-16 Suspense Error Analysis.xlsx").Activate'
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Detail").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Dim lRowAfter As Long
    lRowAfter = Cells(Rows.count, 1).End(xlUp).Row
    Range("AD2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("AD2:AD" & lRowAfter)
    Range("AD2:AD" & lRowAfter).Select
    Range("AE2").Select
    Selection.AutoFill Destination:=Range("AE2:AE" & lRowAfter)
    Range("AE2:AE" & lRowAfter).Select
    Range("AF2").Select
    Selection.AutoFill Destination:=Range("AF2:AF" & lRowAfter)
    Range("AF2:AF" & lRowAfter).Select
    Range("AG2").Select
    Selection.AutoFill Destination:=Range("AG2:AG" & lRowAfter)
    Range("AG2:AG" & lRowAfter).Select
    Range("AH2").Select
    Selection.AutoFill Destination:=Range("AH2:AH" & lRowAfter)
    Range("AH2:AH" & lRowAfter).Select
    Range("AI2").Select
    Selection.AutoFill Destination:=Range("AI2:AI" & lRowAfter)
    Range("AI2:AI" & lRowAfter).Select
    Range("AJ2").Select
    Selection.AutoFill Destination:=Range("AJ2:AJ" & lRowAfter)
    Range("AJ2:AJ" & lRowAfter).Select
    Range("AK2").Select
    Selection.AutoFill Destination:=Range("AK2:AK" & lRowAfter)
    Range("AK2:AK" & lRowAfter).Select
    Range("AL2").Select
    Selection.AutoFill Destination:=Range("AL2:AL" & lRowAfter)
    Range("AL2:AL" & lRowAfter).Select
    Range("AM2").Select
    Selection.AutoFill Destination:=Range("AM2:AM" & lRowAfter)
    Range("AM2:AM" & lRowAfter).Select
    Range("AN2").Select
    Selection.AutoFill Destination:=Range("AN2:AN" & lRowAfter)
    Range("AN2:AN" & lRowAfter).Select
    Range("AO2").Select
    Selection.AutoFill Destination:=Range("AO2:AO" & lRowAfter)
    Range("AO2:AO" & lRowAfter).Select
    Range("AP2").Select
    Selection.AutoFill Destination:=Range("AP2:AP" & lRowAfter)
    Range("AP2:AP" & lRowAfter).Select
    Range("AQ2").Select
    ActiveWindow.SmallScroll ToRight:=3
    Selection.AutoFill Destination:=Range("AQ2:AQ" & lRowAfter)
    Range("AQ2:AQ" & lRowAfter).Select
    Workbooks(fileName).Close SaveChanges:=False
    Application.DisplayAlerts = True




   
  


End Sub

Sub CopyPreviousweekPivotdata()
'
' Macro11 Macro
'

'
   
    Dim SuspenseErrorAnalysisFilePath
    Dim SuspenseErrorAnalysisFileName
    Dim dDate
    Dim lRowNew
    Dim lRowPreviousWeek
    Dim lRowError
    
    SuspenseErrorAnalysisFilePath = DirectoryPath & PreviousweekDate & " Suspense Error Analysis.xlsx"

    'SuspenseErrorAnalysisFilePath = ThisWorkbook.Sheets("Macro").Range("B2")
    SuspenseErrorAnalysisFileName = PreviousweekDate & " Suspense Error Analysis.xlsx"

    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate
    Columns("A:B").Select
    Selection.ClearContents
    dDate = Format(DateAdd("d", -9, Now()), "YYYY-MM-DD")
    Workbooks.Open (SuspenseErrorAnalysisFilePath)
    Worksheets("Week to Week Compare").Activate
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row
    Range("C2:D" & lRowNew).Select
    Selection.Copy
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate
    Range("A2").Select
    ActiveSheet.Paste
    
    

    
    'Workbooks(SuspenseErrorAnalysisFileName).Close SaveChanges:=False
   ' Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Close SaveChanges:=True


End Sub
Sub RefreshSuspenseData()
'
' Macro4 Macro
'
    Dim lRowNew
    Dim lRowPreviousWeek
    Dim lRowError
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Trending Counts").Activate
    Dim lRow As Long
    Dim lRow1 As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    Sheets("All errors by Age").Select
    ActiveWorkbook.RefreshAll
    'lRow1 = Cells(Rows.Count, 9).End(xlUp).Row
    'Range("I" & lRow1).Select
    'Selection.Copy
    'Worksheets("Trending Counts").Activate
    
    'Range("B" & lRow).Select
    'ActiveSheet.Paste
    'Range("C" & lRow - 1).Select
   ' Application.CutCopyMode = False
    'Selection.AutoFill Destination:=Range("C" & lRow - 1 & ":" & "C" & lRow), Type:=xlFillDefault
  
    Worksheets("Week to Week Compare").Activate
    
     Range("B1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY()-7,""MM/DD/yyyy"")"
    Range("B1").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
End Sub




Sub FormatCurrentFileandSave()
'
' Macro9 Macro
'

'
    Dim lRowW
    Dim lRowNew
    Dim lRowPreviousWeek
    Dim lRowError
    Dim i, j, res, count
     count = 0
     
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate

    Range("D1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
    
    
    Range("D1").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

   
    Range("I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("H2").Select
    ActiveSheet.Paste
    Range("I2").Select
    Application.CutCopyMode = False
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
  
    Range("I2").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

 Range("H2").Select
 Range("H2").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    
    Dim SuspenseErrorAnalysisFileName
    SuspenseErrorAnalysisFileName = PreviousweekDate & " Suspense Error Analysis.xlsx"

   
    Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate
    lRowW = Cells(Rows.count, 9).End(xlUp).Row
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row
    
    Range("L3").Select
    
    For j = 3 To lRowNew
    For i = 3 To lRowW
    
    
    
    If Trim(Range("C" & j).Value) = ":" And Trim(Range("G" & i).Value) = ":" Then
    res = Range("D" & j).Value
    
    Range("L" & i).Value = res
    count = 1
    Exit For
    End If
    
    If Range("C" & j).Value = "Grand Total" Or Range("C" & j).Value = "" Then
        
    Exit For
    
    Else
    
    'If (InStr(Range("G" & i).Value, Range("C" & j).Value) > 0) Then
    
    
    
  ' MsgBox (Range("C" & j).Value)
     If ((InStr(Range("C" & j).Value, Range("G" & i).Value) > 0) Or (InStr(Range("G" & i).Value, Range("C" & j).Value) > 0)) And (Trim(Range("C" & j).Value) <> ":" And Trim(Range("G" & i).Value) <> ":") Then
   
    
     res = Range("D" & j).Value
     Range("L" & i).Value = res
     
       count = 1
     Exit For
    
     
        
   
     End If
     End If
     
        Next
        
        If count = 0 And Range("C" & j).Value <> "Grand Total" And Range("D" & j).Value <> "" Then
        Range("G" & lRowW - 1).EntireRow.Insert
        
        Range("G" & lRowW - 1).Value = Range("C" & j).Value
        Range("I" & lRowW - 1).Value = Range("D" & j).Value

        Range("H" & lRowW - 1).Value = 0
        Range("L" & lRowW - 1).Value = Range("D" & j).Value
        
        Range("J" & lRowW - 2).Select
        Selection.AutoFill Destination:=Range("J" & lRowW - 2 & ":J" & lRowW - 1), Type:=xlFillDefault
   
        
        End If
        
        
        
       count = 0
    
Next

lRowW = Cells(Rows.count, 8).End(xlUp).Row

    Range("M3").Select
    
   Range("M3:M" & lRowW - 1).Formula = "=IF(L3=" & """""" & ",0,L3)"
     
     

     


   
   ' ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-5]&""*"",C[-9]:C[-8],2,0),0)"
    
    
    
    
    'Range("L3").Select
    'Selection.AutoFill Destination:=Range("L3:L" & lRowW - 1), Type:=xlFillDefault
    Range("M3:M" & lRowW - 1).Select
   
    Application.CutCopyMode = False
    Selection.Copy
    Range("I3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:M").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
    Worksheets("Week to Week Compare").Activate
    
    
    Range("H" & lRowW + 1).Select
   Range("H" & lRowW + 1).Formula = "=Sum(H3" & ":H" & lRowW - 1 & ")"
    Range("H" & lRowW + 1).Select
    Selection.Copy
    Range("H" & lRowW).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
        
    Range("I" & lRowW + 1).Select
   Range("I" & lRowW + 1).Formula = "=Sum(I3" & ":I" & lRowW - 1 & ")"
    Range("I" & lRowW + 1).Select
    Selection.Copy
    Range("I" & lRowW).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
         
        
    Range("I" & lRowW + 1).ClearContents
         
        Range("H" & lRowW + 1).ClearContents
        

   ' lRowPreviousWeek = Cells(Rows.Count, 2).End(xlUp).Row
    'lRowError = Cells(Rows.Count, 7).End(xlUp).Row
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row
    'Range("D" & lRowNew).Select
    'Selection.Copy
    
    'Range("I" & lRowError).Select
    'ActiveSheet.Paste
   ' Range("B" & lRowPreviousWeek).Select
   ' Application.CutCopyMode = False
    'Selection.Copy
    'Range("H" & lRowError).Select
    'ActiveSheet.Paste
    
    Worksheets("Week to Week Compare").Activate

    Range("D" & lRowNew).Select
    Selection.Copy
    Worksheets("Trending Counts").Activate
    Dim lRow As Long
    Dim lRow1 As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
   
     
    'Selection.Copy
   
    
    Range("B" & lRow).Select
    ActiveSheet.Paste
    Range("C" & lRow - 1).Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("C" & lRow - 1 & ":" & "C" & lRow), Type:=xlFillDefault
    
    Workbooks(SuspenseErrorAnalysisFileName).Close SaveChanges:=False
   Workbooks(Format(Now(), "YYYY-MM-DD") & " Suspense Error Analysis.xlsx").Close SaveChanges:=True
     

End Sub














Sub DownloadFromSharepoint()



Dim myURL As String
Dim strFilePath As String
Dim strFileExists As String
Dim wb_macro As Workbook

Set wb_macro = ThisWorkbook

'On Error GoTo Failure



'Get download path
'strFilePath = wb_macro.Sheets("Home").Range("_rng_DownloadPath").Value
'get sharepoint path
strFilePath = "C:\Users\divyajain24\Desktop\Automation Plan\Suspense And CEW"
myURL = "https://amedeloitte.sharepoint.com/sites/FreseniusRPAteam/Shared%20Documents/"



Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
WinHttpReq.Open "GET", myURL, False
WinHttpReq.send


Dim oStream As Object
If WinHttpReq.Status = 200 Then
Set oStream = CreateObject("ADODB.Stream")
oStream.Open
oStream.Type = 1
oStream.Write WinHttpReq.responseBody
oStream.SaveToFile (strFilePath & "\Test_" & Format(Now, "MMDDYYY_HHMMSS") & ".xlsx") '1 = no overwrite, 2 = overwrite
oStream.Close
End If
errorHandler:

    Dim Subject, Body, olApp, olMailItm
    Subject = "Exception Notification mail"
    
    Body = "Hi Team,<br><br> File generation could not be completed, due to the below error : <br> Error detail : " + Err.Description + " <br><br>Thanks,<br>excel macro"
    Set olApp = CreateObject("Outlook.Application")
    Set olMailItm = olApp.CreateItem(0)
    olMailItm.To = "Divyajain24@deloitte.com"
    olMailItm.Subject = Subject
    olMailItm.BodyFormat = 2
    ' 1 – text format of an email, 2 -  HTML format
    olMailItm.HTMLBody = Body
    On Error GoTo ErrorHandler2
    olMailItm.send
    
ErrorHandler2:


'Failure:



'wb_macro.Sheets("Home").Range("E7").Value = Err.Description

End Sub







Sub CEW()



Dim oFSO As Object
Dim oFolder As Object
Dim oFile As Object
Dim i As Integer

 DirectoryPath = ThisWorkbook.Sheets("Macro").Range("B3")
 
Set oFSO = CreateObject("Scripting.FileSystemObject")

 
Set oFolder = oFSO.GetFolder(DirectoryPath)
 
For Each oFile In oFolder.Files


    If InStr(oFile.Name, ".csv") > 0 Then
         fileName = oFile.Name
        
      End If
   
    Next oFile

 
 
End Sub

Sub PasteFromlastweek()
'
' Macro3 Macro
'

'
    'dDate = Format(DateAdd("d", -9, Now()), "YYYY-MM-DD")'
    
    PreviousweekDate = Format(DateAdd("ww", -1, Format(Now, "yyyy-MM-DD")), "yyyy-MM-DD")
    Application.DisplayAlerts = False
    'Dim SuspenseErrorAnalysisFilePath  As String
    'Dim DirectoryPath  As String
    'Dim Path As String'
    Dim CEWAnalysisFilePath
    Dim wb
    Dim lRowTC
    
    'Workbooks("C:\Users\divyajain24\Desktop\Suspense Reports\Macro.xlsm").Activate
    DirectoryPath = ThisWorkbook.Sheets("Macro").Range("B3")
    Path = DirectoryPath & fileName
    
    CEWAnalysisFilePath = DirectoryPath & PreviousweekDate & " CEWL Error Analysis.xlsx"
    
    Set wb = Workbooks.Open(CEWAnalysisFilePath)
    wb.SaveAs fileName:=DirectoryPath & Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx", FileFormat:=xlOpenXMLStrictWorkbook
    
    
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Trending Counts").Activate
    lRowTC = Cells(Rows.count, 1).End(xlUp).Row
    
    Range("A" & lRowTC + 1).Select
    
    
    Dim K As String
    K = Date
    K = Format(K, "MM/DD/yy")
    ActiveCell.FormulaR1C1 = K
    
    Range("A" & lRowTC + 1).Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    


    
    
    Worksheets("Detail").Activate
    If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
    Dim lRow As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    ActiveSheet.Range("A3:BN" & lRow).Select
    
    Application.CutCopyMode = False
    Selection.ClearContents
    Workbooks.Open (Path)
    Workbooks(fileName).Activate
    ActiveWindow.LargeScroll ToRight:=-1
    Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    'Windows("2021-11-16 Suspense Error Analysis.xlsx").Activate'
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Detail").Activate
    Range("A1").Select
    ActiveSheet.Paste
    Dim lRowAfter As Long
    lRowAfter = Cells(Rows.count, 1).End(xlUp).Row
    Range("BG2").Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("BG2:BG" & lRowAfter)
    Range("BG2:BG" & lRowAfter).Select
    Range("BH2").Select
    Selection.AutoFill Destination:=Range("BH2:BH" & lRowAfter)
    Range("BH2:BH" & lRowAfter).Select
    Range("BI2").Select
    Selection.AutoFill Destination:=Range("BI2:BI" & lRowAfter)
    Range("BI2:BI" & lRowAfter).Select
    Range("BJ2").Select
    Selection.AutoFill Destination:=Range("BJ2:BJ" & lRowAfter)
    Range("BJ2:BJ" & lRowAfter).Select
    Range("BK2").Select
    Selection.AutoFill Destination:=Range("BK2:BK" & lRowAfter)
    Range("BK2:BK" & lRowAfter).Select
    Range("BL2").Select
    Selection.AutoFill Destination:=Range("BL2:BL" & lRowAfter)
    Range("BL2:BL" & lRowAfter).Select
    Range("BM2").Select
    Selection.AutoFill Destination:=Range("BM2:BM" & lRowAfter)
    Range("BM2:BM" & lRowAfter).Select
    
    Range("BN2").Select
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
    Range("BN2").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    Workbooks(fileName).Close SaveChanges:=False
    

End Sub





Sub Comparefrompreviousweek()
'
' Macro9 Macro
'

'
    Dim CEWAnalysisFilePath
    Dim dDate
    Dim lRowNew
    Dim lRowPreviousWeek
    Dim lRowError
    DirectoryPath = ThisWorkbook.Sheets("Macro").Range("B3")
    PreviousweekDate = Format(DateAdd("ww", -1, Format(Now, "yyyy-MM-DD")), "yyyy-MM-DD")
    CEWAnalysisFilePath = DirectoryPath & PreviousweekDate & " CEWL Error Analysis.xlsx"
    
    

    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    
    Worksheets("Week to Week Compare").Activate
    
    Columns("A:B").Select
    Selection.ClearContents
    dDate = Format(DateAdd("d", -9, Now()), "YYYY-MM-DD")
    Workbooks.Open (CEWAnalysisFilePath)
    Worksheets("Week to Week Compare").Activate
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row
    Range("C2:D" & lRowNew).Select
    Selection.Copy
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate
    Range("A2").Select
    ActiveSheet.Paste
    
    
    
    
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Trending Counts").Activate
    Dim lRow As Long
    Dim lRow1
    Dim i, j
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    Sheets("Aging Bar").Select
    ActiveWorkbook.RefreshAll
    'lRow1 = Cells(Rows.Count, 9).End(xlUp).Row
    'Range("I" & lRow1).Select
    'Selection.Copy
    'Worksheets("Trending Counts").Activate
    
    'Range("B" & lRow).Select
    'ActiveSheet.Paste
    Range("C" & lRow - 1).Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("C" & lRow - 1 & ":" & "C" & lRow), Type:=xlFillDefault
    
    
    
    Worksheets("Week to Week Compare").Activate
    lRowPreviousWeek = Cells(Rows.count, 2).End(xlUp).Row
    lRowError = Cells(Rows.count, 7).End(xlUp).Row
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row


    Range("H" & lRowError + 1).Select
   Range("H" & lRowError + 1).Formula = "=Sum(H3" & ":H" & lRowError - 1 & ")"
    Range("H" & lRowError + 1).Select
    Selection.Copy
    Range("H" & lRowError).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("H" & lRowError + 1).ClearContents
        
  Range("G" & lRowError + 1).Select
   Range("G" & lRowError + 1).Formula = "=Sum(G3" & ":G" & lRowError - 1 & ")"
    Range("G" & lRowError + 1).Select
    Selection.Copy
    Range("G" & lRowError).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("G" & lRowError + 1).ClearContents
    
    
    
     Range("B1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY()-7,""MM/DD/yyyy"")"
    
    Range("B1").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    
    'Workbooks(PreviousweekDate & " CEWL Error Analysis.xlsx").Close SaveChanges:=False
    'Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Close SaveChanges:=True
   
    'Kill DirectoryPath & FileName

   
    Dim lRowW
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate

    Range("D1").Select
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
    
    Range("D1").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Dim res
   
    Range("H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("G2").Select
    ActiveSheet.Paste
    Range("G2").Select
    'Application.CutCopyMode = False
    'Selection.ClearContents
    'ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Activate
    Worksheets("Week to Week Compare").Activate
    lRowW = Cells(Rows.count, 8).End(xlUp).Row
    
    Range("L3").Select
    
     lRowNew = Cells(Rows.count, 3).End(xlUp).Row
    Dim count
    
    'For j = 3 To lRowNew
    'For i = 3 To lRowW
    
    
    
   ' If Trim(Range("C" & j).Value) = ":" And Trim(Range("F" & i).Value) = ":" Then
    'res = Range("D" & j).Value
    
    'Range("L" & i).Value = res
    'count = 1
    'Exit For
    'End If
    
   ' If Range("C" & j).Value = "Grand Total" Or Range("C" & j).Value = "" Then
        
   ' Exit For
    
    'Else
    
    'If (InStr(Range("G" & i).Value, Range("C" & j).Value) > 0) Then
    
  ' MsgBox (Range("C" & j).Value)
     'If ((InStr(Range("C" & j).Value, Range("F" & i).Value) > 0) Or (InStr(Range("F" & i).Value, Range("C" & j).Value) > 0)) And (Trim(Range("C" & j).Value) <> ":" And Trim(Range("F" & i).Value) <> ":") Then
    'If Trim(Range("G" & i).Value) <> ":" And Trim(Range("C" & j).Value) <> ":" Then
    
     'res = Range("D" & j).Value
     'Range("L" & i).Value = res
     
      ' count = 1
     'Exit For
    
     
        
   
    ' End If
     'End If
     
       ' Next
        
        'If count = 0 And Range("C" & j).Value <> "Grand Total" Then
        'Range("L" & i).Value = 0
        'End If
        
        
        
      ' count = 0
    
'Next

For j = 3 To lRowNew
    For i = 3 To lRowW
    
    
    
    If Trim(Range("C" & j).Value) = ":" And Trim(Range("F" & i).Value) = ":" Then
    res = Range("D" & j).Value
    
    Range("L" & i).Value = res
    count = 1
    Exit For
    End If
    
    If Range("C" & j).Value = "Grand Total" Or Range("C" & j).Value = "" Then
        
    Exit For
    
    Else
    
    'If (InStr(Range("G" & i).Value, Range("C" & j).Value) > 0) Then
    
    
    
  ' MsgBox (Range("C" & j).Value)
     If ((InStr(Range("C" & j).Value, Range("F" & i).Value) > 0) Or (InStr(Range("F" & i).Value, Range("C" & j).Value) > 0)) And (Trim(Range("C" & j).Value) <> ":" And Trim(Range("F" & i).Value) <> ":") Then
   
    
     res = Range("D" & j).Value
     Range("L" & i).Value = res
     
       count = 1
     Exit For
    
     
        
   
     End If
     End If
     
        Next
        
        If count = 0 And Range("C" & j).Value <> "Grand Total" And Range("D" & j).Value <> "" Then
        Range("F" & lRowW - 1).EntireRow.Insert
        
        Range("F" & lRowW - 1).Value = Range("C" & j).Value
        Range("H" & lRowW - 1).Value = Range("D" & j).Value

        Range("G" & lRowW - 1).Value = 0
        Range("L" & lRowW - 1).Value = Range("D" & j).Value
        
        Range("I" & lRowW - 2).Select
        Selection.AutoFill Destination:=Range("I" & lRowW - 2 & ":I" & lRowW - 1), Type:=xlFillDefault
   
        
        End If
        
        
        
       count = 0
    
Next

lRowW = Cells(Rows.count, 8).End(xlUp).Row

    Range("M3").Select
    
   Range("M3:M" & lRowW - 1).Formula = "=IF(L3=" & """""" & ",0,L3)"

    Range("M3").Select
    Range("M3:M" & lRowW - 1).Formula = "=IF(L3=" & """""" & ",0,L3)"
   
    'ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-6]&""*"",C[-9]:C[-8],2,0),0)"
    
  
    
    

    Range("M3").Select
    'Selection.AutoFill Destination:=Range("L3:L" & lRowW - 1), Type:=xlFillDefault
    Range("M3:M" & lRowW - 1).Select
    Range("M3:M" & lRowW - 1).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("H3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Columns("L:M").Select
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
     Range("H2").Select
     Application.CutCopyMode = False
    Selection.ClearContents
    ActiveCell.FormulaR1C1 = "=TEXT(TODAY(),""MM/DD/yyyy"")"
    Range("H2").Copy
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

    Range("G2").Select

    
    Range("G2").Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        Range("B" & lRowPreviousWeek).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("G" & lRowError).Select
    ActiveSheet.Paste
    'Range("D" & lRowNew).Select
    'Selection.Copy
    
   ' Range("H" & lRowError).Select
    'ActiveSheet.Paste
    
    
    Range("H" & lRowError + 1).Select
    Range("H" & lRowError + 1).Formula = "=Sum(H3" & ":H" & lRowError - 1 & ")"
    Range("H" & lRowError + 1).Select
    Selection.Copy
    Range("H" & lRowError).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Range("H" & lRowError + 1).ClearContents
    
    
    Worksheets("Week to Week Compare").Activate
    lRowPreviousWeek = Cells(Rows.count, 2).End(xlUp).Row
    lRowError = Cells(Rows.count, 7).End(xlUp).Row
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row

    
    Range("D" & lRowNew).Select
    Selection.Copy
    
    Worksheets("Trending Counts").Activate
    Range("B" & lRow).Select
    ActiveSheet.Paste
    Range("C" & lRow - 1).Select
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("C" & lRow - 1 & ":" & "C" & lRow), Type:=xlFillDefault

    Workbooks(PreviousweekDate & " CEWL Error Analysis.xlsx").Close SaveChanges:=False
    Workbooks(Format(Now(), "YYYY-MM-DD") & " CEWL Error Analysis.xlsx").Close SaveChanges:=True
   
    Kill DirectoryPath & fileName

End Sub



Sub test()
Dim lRowW
Dim lRowNew, j, i, res
Dim count

 Workbooks.Open ("C:\Users\divyajain24\Desktop\Suspense Reports\2022-05-03 CEWL Error Analysis.xlsx")
  lRowW = Cells(Rows.count, 8).End(xlUp).Row
    lRowNew = Cells(Rows.count, 3).End(xlUp).Row

For j = 3 To lRowNew
    For i = 3 To lRowW
    
    
    
    If Trim(Range("C" & j).Value) = ":" And Trim(Range("F" & i).Value) = ":" Then
    res = Range("D" & j).Value
    
    Range("L" & i).Value = res
    count = 1
    Exit For
    End If
    
    If Range("C" & j).Value = "Grand Total" Or Range("C" & j).Value = "" Then
        
    Exit For
    
    Else
    
    'If (InStr(Range("G" & i).Value, Range("C" & j).Value) > 0) Then
    
  ' MsgBox (Range("C" & j).Value)
     If ((InStr(Range("C" & j).Value, Range("F" & i).Value) > 0) Or (InStr(Range("F" & i).Value, Range("C" & j).Value) > 0)) And (Trim(Range("C" & j).Value) <> ":" And Trim(Range("F" & i).Value) <> ":") Then
    'If Trim(Range("G" & i).Value) <> ":" And Trim(Range("C" & j).Value) <> ":" Then
    
     res = Range("D" & j).Value
     Range("L" & i).Value = res
     
       count = 1
     Exit For
    
     
        
   
     End If
     End If
     
        Next
        
        If count = 0 And Range("C" & j).Value <> "Grand Total" Then
        Range("L" & i).Value = 0
        End If
        
        
        
       count = 0
    
Next

    Range("M3").Select
   Range("M3:M" & lRowW - 1).Formula = "=IF(L3=" & """""" & ",0,L3)"
     
End Sub


