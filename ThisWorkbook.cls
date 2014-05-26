VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim MekkoReceiverPort As Long
Dim MekkoSenderPort As Long
Dim DebugTrace As Boolean

'----------------------------------
'    Save this file interactive
'----------------------------------
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

'----------------------------------
'             Debugging
'----------------------------------
Private Sub DebugTraceMsg(msg As String)
    If DebugTrace Then
        If msg = "Activated" Or msg = "SheetChange" Then
            MsgBox "__MekkoExcel__.xlsm:   " & msg & " :  MekkoReceiverPort=" & MekkoReceiverPort
        Else
            MsgBox "__MekkoExcel__.xlsm:   " & msg
        End If
    End If
End Sub

'----------------------------------
'             Interactive debugging
'----------------------------------
Sub SetTraceByMekko()
    If DebugTrace Then
        If MsgBox("Set DebugTrace to Off?", vbYesNo, "Confirm") = vbYes Then DebugTrace = False
    Else
        If MsgBox("Set DebugTrace to On?", vbYesNo, "Confirm") = vbYes Then DebugTrace = True
    End If
End Sub

Sub TraceByMekko(flag As Integer)
    If flag = 0 Then
        DebugTrace = False
    Else
        DebugTrace = True
    End If
End Sub

'----------------------------------
'      Interactive Status
'----------------------------------
Sub StatusByMekko()
    Dim msg As String
    Dim oldMekkoReceiverPort As Long
    Dim oldMekkoSenderPort As Long
    
    oldMekkoReceiverPort = MekkoReceiverPort
    oldMekkoSenderPort = MekkoSenderPort
    CheckMekkoApp
    
    msg = "MekkoReceiverPort=" & MekkoReceiverPort
    If Not oldMekkoReceiverPort = MekkoReceiverPort Then
        msg = msg & " (was " & oldMekkoReceiverPort & ")"
    End If
    
    msg = msg & "   MekkoSenderPort=" & MekkoSenderPort
    If Not oldMekkoSenderPort = MekkoSenderPort Then
        msg = msg & "(was " & oldMekkoSenderPort & ")"
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

Sub TestByMekko()
    '----------------------------------
    '          Test
    '----------------------------------
End Sub

'----------------------------------
'      Checks: is Mekko application started?
'----------------------------------
Private Function IsMekkoActive() As Boolean
    If InStr(MacScript("do shell script ""ps -ax | grep MGEditor"""), "MGEditor.app/Contents/MacOS/MGEditor") = 0 Then
        IsMekkoActive = False
    Else
        IsMekkoActive = True
    End If
End Function

Sub CheckMekkoApp()
    If Not IsMekkoActive Then
        MekkoReceiverPort = 0
        MekkoSenderPort = 0
    End If
End Sub

'----------------------------------
'      Save Mekko communication ports
'----------------------------------
Sub SetMekkoReceiverPort(port As Long)
    MekkoReceiverPort = port
    If port < 1024 Then MekkoReceiverPort = 0
End Sub

Sub SetMekkoSenderPort(port As Long)
    MekkoSenderPort = port
    If port < 1024 Then MekkoSenderPort = 0
End Sub

'----------------------------------
'      Send Information to Mekko over TCP/IP
'----------------------------------
Private Sub SendCmdToMekko(text As String)
    Dim controlCode As String
    controlCode = Chr(8)
    SendTextToMekko (controlCode & text)
End Sub

Private Sub SendTextToMekko(text As String)
    If MekkoReceiverPort > 1024 Then
        Dim msg As String
        Dim cmd As String
        
        msg = Replace(text, "$", Chr(2))
        msg = Replace(msg, "#", Chr(3))
        msg = Replace(msg, "'", Chr(5))
        msg = Replace(msg, """", Chr(6))
        msg = Replace(msg, "\", Chr(7))
        
        cmd = "do shell script ""echo '" & msg & Chr(4) & "' | nc 127.0.0.1 " & MekkoReceiverPort & """"
        
        On Error GoTo ErrorHandling
            MacScript (cmd)
        GoTo EndSub
        
ErrorHandling:
        CheckMekkoApp
EndSub:
End If

End Sub

'----------------------------------
'     SetRangeValueByMekko
'           (rowBeg, colBeg) - top left corner of Range
'           (rowCount, colCount) - Dimansions of Range array
'----------------------------------
Sub SetRangeValueByMekko(ByVal rowBegVal As String, ByVal colBegVal As String, ByVal rowCountVal As String, ByVal colCountVal As String)
    CheckMekkoApp
    If Not IsMekkoActive Or MekkoSenderPort = 0 Then
        SendCmdToMekko ("Error: MekkoSenderPort is not active")
        GoTo EndSub
    End If
    
    Dim rowBeg As Long, rowEnd As Long, rowCount As Long
    Dim colBeg As Long, colEnd As Long, colCount As Long
    
    rowBeg = rowBegVal
    rowCount = rowCountVal
    rowEnd = rowBeg + rowCount - 1
    colBeg = colBegVal
    colCount = colCountVal
    colEnd = colBeg + colCount - 1

    Dim cmd As String, data As String
    cmd = "do shell script ""echo '" & Chr(4) & "' | nc 127.0.0.1 " & MekkoSenderPort & """"
    On Error GoTo ErrorHandling
    data = MacScript(cmd)
    
    Dim strArray() As String, delim As String
    delim = Chr(9)
    strArray = Split(data, delim)
    
    If 2 * rowCount * colCount > (UBound(strArray) - LBound(strArray) + 1) Then
        SendCmdToMekko ("Error: Wrong data lenth")
        GoTo EndSub
    End If

    Dim valArray() As Variant
    ReDim valArray(1 To rowCount, 1 To colCount)
    ReDim fmtArray(1 To rowCount, 1 To colCount)
    
    Application.ScreenUpdating = False
    index = LBound(strArray)
    For i = 1 To rowCount
        For j = 1 To colCount
            valArray(i, j) = strArray(index)
            index = index + 1
            fmtArray(i, j) = strArray(index)
            index = index + 1
        Next j
    Next i
        
    Dim TheRange As Range
    Set TheRange = Range(Cells(rowBeg, colBeg), Cells(rowEnd, colEnd))
    TheRange.Value = valArray
    TheRange.NumberFormat = fmtArray
    
    Application.ScreenUpdating = True
    GoTo EndSub
    
ErrorHandling:
    SendCmdToMekko ("Error: Wrong cmd=" & cmd)
EndSub:
End Sub

'----------------------------------
'    Mekko commands called by AppleScript
'----------------------------------
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

'----------------------------------
'    Excel event: open this workbook
'----------------------------------
Private Sub Workbook_Open()
    MekkoReceiverPort = 0
    MekkoSenderPort = 0
    
    DebugTrace = False
    HideByMekko
    DebugTraceMsg ("Open")
    If Not IsMekkoActive Then
        DebugTraceMsg "Add Workbook"
        Workbooks.Add
    End If
End Sub

'-----------------------------------------------------
'    Excel event: workbook closed
'-----------------------------------------------------
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    DebugTraceMsg ("BeforeClose")
    If IsWorkbookVisible Then
        DebugTraceMsg ("Hide")
        HideByMekko
        CheckMekkoApp
        If MekkoReceiverPort > 1024 Then
            SendCmdToMekko ("close")
        End If
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

'-----------------------------------------------------
'    Excel event: worksheet changed, send everything to Mekko
'-----------------------------------------------------
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    DebugTraceMsg ("SheetChange")
    If MekkoReceiverPort > 1024 Then

        Dim KeyCells As Range
        Set KeyCells = ActiveSheet.UsedRange
    
        Application.CalculateFull
        
        Dim msg As String, delim As String
        delim = Chr(9)
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
