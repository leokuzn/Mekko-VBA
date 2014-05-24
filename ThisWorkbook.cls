VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim CommunicationPort As Long
Dim ControlPort As Long
Dim DebugTrace As Boolean

Private Sub DebugTraceMsg(msg As String)
    If DebugTrace Then
        If msg = "Activated" Or msg = "SheetChange" Then
            MsgBox "__MekkoExcel__.xlsm:   " & msg & " :  CommunicationPort=" & CommunicationPort
        Else
            MsgBox "__MekkoExcel__.xlsm:   " & msg
        End If
    End If
End Sub

Private Function IsMekkoActive() As Boolean
    If InStr(MacScript("do shell script ""ps -ax | grep MGEditor"""), "MGEditor.app/Contents/MacOS/MGEditor") = 0 Then
        IsMekkoActive = False
    Else
        IsMekkoActive = True
    End If
End Function

Sub CheckMekkoApp()
    If Not IsMekkoActive Then
        CommunicationPort = 0
    End If
End Sub

Sub SetTraceByMekko()
    If DebugTrace Then
        If MsgBox("Set DebugTrace to Off?", vbYesNo, "Confirm") = vbYes Then DebugTrace = False
    Else
        If MsgBox("Set DebugTrace to On?", vbYesNo, "Confirm") = vbYes Then DebugTrace = True
    End If
End Sub

Sub StartUpdateByMekko()
    Workbooks("__MekkoExcel__.xlsm").Worksheets("Sheet1").Activate
End Sub

Sub EndUpdateByMekko()
    Workbooks("__MekkoExcel__.xlsm").Worksheets("Sheet1").Activate
    Range("A1").Select
End Sub

Sub SetCellByMekko(ByVal cAddr As String, ByVal cVal As String, ByVal cFormat As String, ByVal cPrefix As String, ByVal cFormula As String)
    Addr = Application.ConvertFormula(Formula:="=" & cAddr, fromReferenceStyle:=xlR1C1, toReferenceStyle:=xlA1)
    With Range(Addr)
        .Clear
        If Not cFormula = "" Then
            .Formula = cFormula
        Else
            If Not cPrefix = "" Then cVal = cPrefix & cVal
            .Value = cVal
        End If
        If cFormat = "" Then
            .NumberFormat = "General"
        Else
            .NumberFormat = cFormat
        End If
        .Select
    End With
End Sub

Sub TestByMekko()
    Workbooks("__MekkoExcel__.xlsm").Worksheets("Sheet1").Activate
    Dim rng As Range
    Set rng = Selection
    Selection.Clear
    
    'Application.ReferenceStyle = xlA1
    
    'Dim cmd As String
    'cmd = "do shell script ""nc -v -l 1234"""
    'msg = MacScript(cmd)
    'MsgBox msg
End Sub

Sub StatusByMekko()
    Dim msg As String
    Dim oldCommunicationPort As Long
    oldCommunicationPort = CommunicationPort
    CheckMekkoApp
    
    msg = "CommunicationPort=" & CommunicationPort
    If Not oldCommunicationPort = CommunicationPort Then
        msg = msg & " (was " & oldCommunicationPort & ")"
    End If
    msg = msg & vbCrLf
    
    Dim aName As String
    aName = ""
    On Error GoTo NoActive
    aName = Application.ActiveWorkbook.Name
    
NoActive:
    Dim wb As Workbook
    For Each wb In Workbooks
        msg = msg & vbCrLf & wb.Name
        If wb.Name = aName Then
            msg = msg & " <---- Active"
            If wb.Windows(1).Visible Then msg = msg & ", Visible"
        Else
            If wb.Windows(1).Visible Then msg = msg & "<----  Visible"
        End If
    Next wb
    MsgBox msg

End Sub

Sub SetCommunicationPortByMekko(port As Long)
    CommunicationPort = port
    If port < 1024 Then CommunicationPort = 0
End Sub

Private Sub SendTextToMekko(text As String)
    If CommunicationPort > 1024 Then
        Dim msg As String
        Dim cmd As String
        
        msg = Replace(text, "$", Chr(2))
        msg = Replace(msg, "#", Chr(3))
        msg = Replace(msg, "'", Chr(5))
        msg = Replace(msg, """", Chr(6))
        msg = Replace(msg, "\", Chr(7))
        
        cmd = "do shell script ""echo '" & msg & Chr(4) & "' | nc 127.0.0.1 " & CommunicationPort & """"
        
        On Error GoTo ErrorHandling
            MacScript (cmd)
        GoTo EndSub
        
ErrorHandling:
        CheckMekkoApp
EndSub:
End If

End Sub

Sub TraceByMekko(flag As Integer)
    If flag = 0 Then
        DebugTrace = False
    Else
        DebugTrace = True
    End If
End Sub

Sub CloseByMekko()
    Dim wb As Workbook
    Dim doQuit As Boolean
    
    doQuit = True
    For Each wb In Workbooks
        If wb.Name <> "__MekkoExcel__.xlsm" Then
            doQuit = False
        Else
            HideByMekko
        End If
    Next
    
    If doQuit Then Application.Quit
End Sub

Private Function IsWorkbookVisible() As Boolean
    IsWorkbookVisible = Workbooks("__MekkoExcel__.xlsm").Windows(1).Visible
End Function

Sub HideByMekko()
    Workbooks("__MekkoExcel__.xlsm").Windows(1).Visible = False
End Sub

Sub UnhideByMekko()
    Workbooks("__MekkoExcel__.xlsm").Windows(1).Visible = True
End Sub

Sub CleanupByMekko()
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.Cells.Clear
        'ws.Range("A1").Select
    Next ws
End Sub

Sub SaveByMekko()
    Dim alert As Boolean
    alert = Application.DisplayAlerts
    Application.DisplayAlerts = False
    If MsgBox("Clean up all worksheets ?", vbYesNo, "Confirm") = vbYes Then
        CleanupByMekko
    End If

    If IsWorkbookVisible Then
        DebugTraceMsg ("Hide and Save")
        HideByMekko
        ThisWorkbook.Save
        UnhideByMekko
    Else
        DebugTraceMsg ("Save")
        ThisWorkbook.Save
    End If
    Application.DisplayAlerts = alert
End Sub

Private Sub Workbook_Open()
    CommunicationPort = 0
    ControlPort = 21954
    DebugTrace = False
    HideByMekko
    DebugTraceMsg ("Open")
    If Not IsMekkoActive Then
        DebugTraceMsg "Add Workbook"
        Workbooks.Add
    End If
End Sub

Private Sub Workbook_Activate()
    DebugTraceMsg ("Activated")
End Sub

Private Sub Workbook_Deactivate()
    DebugTraceMsg ("Deactivate")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DebugTraceMsg ("BeforeClose")
    If IsWorkbookVisible Then
        DebugTraceMsg ("Hide")
        HideByMekko
        Cancel = True
    Else
            DebugTraceMsg ("Close")
            Me.Saved = True
    End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    DebugTraceMsg ("BeforeSave: Clean up")
    CleanupByMekko
End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
    DebugTraceMsg ("NewSheet")
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    DebugTraceMsg ("SheetActivate")
End Sub

Private Sub Workbook_SheetDeactivate(ByVal Sh As Object)
    DebugTraceMsg ("SheetDeactivate")
End Sub

Private Sub Workbook_WindowActivate(ByVal Wn As Window)
    DebugTraceMsg ("WindowActivate")
End Sub

Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
    DebugTraceMsg ("WindowDeactivate")
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    DebugTraceMsg ("SheetChange")
    If CommunicationPort > 1024 Then

        Dim KeyCells As Range
        Set KeyCells = ActiveSheet.UsedRange
    
        Application.CalculateFull
        
        Dim delim As String
        delim = Chr(9)
        
        Dim msg As String
        msg = ""
        
        For Each c In ActiveSheet.UsedRange.Cells
            If Not IsEmpty(c.Value) Then
                Addr = c.Address(ReferenceStyle:=xlR1C1)
                cVal = c.Value
                fmt = c.NumberFormat
                If fmt = "General" Then fmt = ""
                prefix = c.PrefixCharacter
                frml = ""
                If c.HasFormula Then frml = c.Formula
                msg = msg & Addr & delim & cVal & delim & fmt & delim & prefix & delim & frml & Chr(10)
            End If
        Next
        SendTextToMekko (msg)
    End If
End Sub

Sub ArrayFillRange()
'   Fill a range by transferring an array
    Dim CellsDown As Long, CellsAcross As Long
    Dim i As Long, j As Long
    Dim StartTime As Double
    Dim TempArray() As Variant
    Dim TheRange As Range
    Dim CurrVal As Long

'   Change these values
    CellsDown = 500
    CellsAcross = 200

    Cells.Clear
'   Record starting time
    StartTime = Timer

'   Redimension temporary array
    ReDim TempArray(1 To CellsDown, 1 To CellsAcross)

'   Set worksheet range
    Set TheRange = Range(Cells(1, 1), Cells(CellsDown, CellsAcross))

'   Fill the temporary array
    CurrVal = 0
    Application.ScreenUpdating = False
    For i = 1 To CellsDown
        For j = 1 To CellsAcross
            TempArray(i, j) = CurrVal
            CurrVal = CurrVal + 1
        Next j
    Next i

'   Transfer temporary array to worksheet
    TheRange.Value = TempArray

'   Display elapsed time
    Application.ScreenUpdating = True
    MsgBox Format(Timer - StartTime, "00.00") & " seconds"
End Sub
