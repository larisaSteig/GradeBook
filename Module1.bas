Attribute VB_Name = "Module1"
Option Explicit
Option Base 1

Public tWb As Workbook, aWb As Workbook, path As String, LastColumn As Integer, ColumnLetter As String, FileName As String

Sub OpenFolder()
Dim NewFolderName As String
Dim folder As String, AnswerMess As VbMsgBoxResult
Dim NewFileName As String
On Error GoTo here

AnswerMess = MsgBox("Please, pick the folder where you want to keep the file: ", vbInformation + vbOKCancel)
If AnswerMess = vbCancel Then
MsgBox "Nothing was selected"
Exit Sub
End If

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
        If .Show = 0 Then
            Exit Sub
            Else:
            folder = .SelectedItems(1)
        End If
End With

NewFolderName = InputBox("Please, Enter a New Folder name: ")
' create a new folder
MkDir folder & "\" & NewFolderName

NewFileName = InputBox("Please, enter the File Name:")

ThisWorkbook.SaveAs FileName:=folder & "\" & NewFolderName & "\" & NewFileName & ".xlsm"

MsgBox NewFileName & " is created and located within " & NewFolderName

here:
End Sub

Sub ImportRoster()
Dim RoasterBook As Workbook
Dim tWb As Workbook
Dim FileToOpen As String

With Application
    .ScreenUpdating = False
    .DisplayAlerts = False
End With


Set tWb = ThisWorkbook

If WorksheetExists2("Roster") = True Then
Worksheets("Roster").Delete
End If

FileToOpen = Application.GetOpenFilename(FileFilter:="Excel Files (*.xls*), *xls*", Title:="Pick Roaster file")

Set RoasterBook = Workbooks.Open(FileToOpen)
On Error Resume Next
ActiveSheet.Copy After:=tWb.Worksheets(Worksheets.Count)
ActiveSheet.Name = "Roster"

RoasterBook.Close

With Application
    .ScreenUpdating = True
    .DisplayAlerts = True
End With

tWb.Worksheets("START PAGE").Activate
MsgBox "Done"
End Sub


Sub CreateSections()
Dim aWb As Workbook, tWb As Workbook
Dim minSection As Integer, maxSection As Integer
Dim nNames As Integer, nSections As Integer
Dim folder As String, k As Integer
Dim N() As String, i As Integer, j As Integer
Dim ColumnLetter As String, path As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False

Set tWb = ThisWorkbook

minSection = WorksheetFunction.Min(tWb.Worksheets("Roster").Columns("C:C"))
maxSection = WorksheetFunction.Max(tWb.Worksheets("Roster").Columns("C:C"))
nNames = WorksheetFunction.CountA(tWb.Worksheets("Roster").Columns("A:A"))
nSections = maxSection - minSection + 1

MsgBox ("Please choose the directory/folder where you'd like to place the files.")

With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
        If .Show = 0 Then
            Exit Sub
            Else:
            folder = .SelectedItems(1)
            tWb.Worksheets("Start Page").Range("A1") = folder
        End If
End With

For i = minSection To maxSection
    k = 0
    ReDim N(1) As String
    For j = 1 To nNames
        If tWb.Worksheets("Roster").Range("C" & j) = i Then
            k = k + 1
            ReDim Preserve N(k)
            N(k) = tWb.Worksheets("Roster").Range("A" & j) & " " & tWb.Worksheets("Roster").Range("B" & j)
        End If
    Next j
    Workbooks.Add
    Set aWb = ActiveWorkbook
    For j = 1 To k
        Range("A" & j + 1) = N(j)
        Range("B" & j + 1) = Right(Range("A" & j + 1), 10)
        Range("A" & j + 1) = Left(Range("A" & j + 1), Len(Range("A" & j + 1)) - 10)
    Next j
    Columns("A").AutoFit
    Columns("B").AutoFit
    aWb.Sheets(1).Range("A1") = "Name"
    aWb.Sheets(1).Range("B1") = "Student ID"
    Dim y As Integer
    'Set aWB = ActiveWorkbook
    For y = 1 To Headings.cbxHomework.ListIndex + 1
        aWb.Sheets(1).Range("C1").Offset(0, y - 1) = "HW " & y
    Next y
    
    Dim LastColumn As Long
     'Ctrl + Shift + End
    LastColumn = aWb.Sheets(1).Cells(1, aWb.Sheets(1).Columns.Count).End(xlToLeft).Column
    For y = 1 To Headings.cbxExams.ListIndex + 1
        aWb.Sheets(1).Cells(1, LastColumn + 1).Offset(0, y - 1) = "Exams " & y
    Next y
    
    LastColumn = aWb.Sheets(1).Cells(1, aWb.Sheets(1).Columns.Count).End(xlToLeft).Column
    For y = 1 To Headings.cbxLabs.ListIndex + 1
        aWb.Sheets(1).Cells(1, LastColumn + 1).Offset(0, y - 1) = "Labs " & y
    Next y
    
    LastColumn = aWb.Sheets(1).Cells(1, aWb.Sheets(1).Columns.Count).End(xlToLeft).Column
    ColumnLetter = Split(Cells(1, LastColumn).Address, "$")(1)
    If i = minSection Then
        aWb.Sheets(1).Range("C1:" & ColumnLetter & 1).Copy tWb.Worksheets("Roster").Range("D1")
'        tWb.Activate
'        Worksheets("Roster").Range("D1").PasteSpecial xlPasteValues
        aWb.Activate
    End If
    
    ActiveWorkbook.SaveAs FileName:= _
        folder & "\File_" & i & ".xlsx"
    aWb.Close
Next i


MsgBox " All files are ready. There are total of: " & nSections
'Unload Headings

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Sub SynchFiles()
Dim folder As String, FileName As String, tWb As Workbook, aWb As Workbook, Name As String, LastColumn As Long, ColumnLetter As String
Dim countNamesinSingleBook As Integer, j As Integer, nNames As Integer, i As Integer, nItems As Integer, idx As Integer
Dim G() As Variant
Dim path As String

With Application
.DisplayAlerts = False
.ScreenUpdating = False
End With

Set tWb = ThisWorkbook
nNames = WorksheetFunction.CountA(tWb.Worksheets("Roster").Columns("A:A"))
nItems = WorksheetFunction.CountA(tWb.Worksheets("Roster").Rows(1)) - 3
ReDim G(nItems) As Variant

path = ThisWorkbook.Sheets(1).Range("A1")
FileName = Dir(path & "\*.xlsx")
' entering one excel file at a time
Do
    Workbooks.Open path & "\" & FileName
    Set aWb = ActiveWorkbook
    LastColumn = aWb.Sheets(1).Cells(1, aWb.Sheets(1).Columns.Count).End(xlToLeft).Column
    ColumnLetter = Split(Cells(1, LastColumn).Address, "$")(1)
    
    countNamesinSingleBook = Application.WorksheetFunction.CountA(aWb.Sheets(1).Columns("A:A"))
    ' find the name in the activebook

    For i = 2 To countNamesinSingleBook
        Name = aWb.Sheets(1).Cells(i, 1)
        
        G = Range("C" & i & ":" & ColumnLetter & i)
        If IsArrayEmpty(G) = False Then
        
            tWb.Worksheets("Roster").Activate
            idx = WorksheetFunction.Match(Name, Range("A2:A" & nNames), 0)
            Range("D" & idx + 1).Select 'This will be ActiveCell
            For j = 1 To nItems
                ActiveCell.Offset(0, j - 1) = G(1, j)
            Next j
            aWb.Activate
        Else:
        End If
       
    Next i
    aWb.Close SaveChanges:=True
    FileName = Dir
Loop Until FileName = ""

With Application
.DisplayAlerts = True
.ScreenUpdating = True
End With

MsgBox " All synchronized!"

End Sub

Sub BackUp()
Dim path As String, TodayDate As String
Dim NowArr As Variant
Application.DisplayAlerts = False
path = ActiveSheet.Range("A1")
NowArr = Split(Now(), " ")
TodayDate = Replace(NowArr(0), "/", "-")
ThisWorkbook.SaveCopyAs path + "\Grade_Manager_Backup_" & TodayDate & ".xlsm"
MsgBox "File is saved here: " & path
Application.DisplayAlerts = True
End Sub

Sub AddAssignments()
'AddAssignment.cbxAddAssignment.RowSource = ThisWorkbook.Worksheets("Path").Range("A3:A10").Value
AddAssignment.Show

End Sub

Function IsArrayEmpty(a As Variant) As Boolean
Dim element As Variant
Dim TotalSum As Integer

TotalSum = 0
For Each element In a
TotalSum = TotalSum + element
Next

If TotalSum = 0 Then
IsArrayEmpty = True
End If

End Function

Sub MainForm()
With frmMainForm
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With
End Sub

Sub Add()

Dim aWb As Workbook, tWb As Workbook
Dim minSection As Integer, maxSection As Integer
Dim nNames As Integer, nSections As Integer
Dim folder As String, k As Integer
Dim N() As String, i As Integer, j As Integer, l As Integer
Dim CategoryName As String, path As String
Dim FileName As String
Dim nItems As Integer, nItemsMain As Integer, LastColumn As Long, ColumnLetter As String

Application.DisplayAlerts = False
Application.ScreenUpdating = False

CategoryName = AddAssignment.cbxAddAssignment.Text

Set tWb = ThisWorkbook
nItemsMain = WorksheetFunction.CountA(tWb.Worksheets("Roster").Rows(1)) - 1

path = tWb.Worksheets("Start Page").Range("A1")
FileName = Dir(path & "\*.xlsx")

' entering one excel file at a time
    Do
        Workbooks.Open path & "\" & FileName
        Set aWb = ActiveWorkbook
        nItems = WorksheetFunction.CountA(Rows(1)) ' count number of items in Row A on FileBook)
        
        For i = 1 To nItems
            j = 0
            If InStr(1, Cells(1, i), CategoryName) > 0 Then
            j = j + 1
                If InStr(1, Cells(1, i + 1), CategoryName) = 0 Then
                    Cells(1, i + 1).EntireColumn.Insert
                    Cells(1, i + 1) = CategoryName & " " & CInt(Right(Cells(1, i), Len(Cells(1, i)) - Len(CategoryName))) + 1
                    l = WorksheetFunction.CountA(aWb.Sheets(1).Rows(1))
                        If l <> nItemsMain Then
                            LastColumn = aWb.Sheets(1).Cells(1, aWb.Sheets(1).Columns.Count).End(xlToLeft).Column
                            ColumnLetter = Split(Cells(1, LastColumn).Address, "$")(1)
                            aWb.Sheets(1).Range("C1:" & ColumnLetter & 1).Copy tWb.Worksheets("Roster").Range("D1")
                            nItemsMain = l
'                            tWb.Activate
'                            Worksheets("Roster").Range("D1").PasteSpecial xlPasteValues
'                            nItemsMain = l
                            'nItemsMain = WorksheetFunction.CountA(tWb.Worksheets("Roster").Rows(1)) - 1
                            aWb.Activate
                        End If
                    Exit For
                End If
            End If
        Next i
        aWb.Close True
        FileName = Dir()
    Loop Until FileName = ""
    
MsgBox "A column for " & CategoryName & " has been added!"

'Unload AddAssignment

Application.DisplayAlerts = True
Application.ScreenUpdating = True
End Sub

Function WorksheetExists2(WorksheetName As String, Optional wb As Workbook) As Boolean
    If wb Is Nothing Then Set wb = ThisWorkbook
    With wb
        On Error Resume Next
        WorksheetExists2 = (.Sheets(WorksheetName).Name = WorksheetName)
        On Error GoTo 0
    End With
End Function
