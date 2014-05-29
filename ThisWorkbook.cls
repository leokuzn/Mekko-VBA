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
Const myName As String = "__MekkoExcel__.xlsb"

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
            MsgBox myName & ":   " & msg & " :  MekkoReceiverPort=" & MekkoReceiverPort
        Else
            MsgBox myName & ":   " & msg
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
        
        If Not text = "" Then
            msg = Replace(text, "$", Chr(2))
            msg = Replace(msg, "#", Chr(3))
            msg = Replace(msg, "'", Chr(5))
            msg = Replace(msg, """", Chr(6))
            msg = Replace(msg, "\", Chr(7))
        End If
        
        cmd = "do shell script ""echo '" & msg & Chr(4) & "' | nc 127.0.0.1 " & MekkoReceiverPort & """"
        
        On Error GoTo ErrorHandling
            MacScript (cmd)
        GoTo EndSub
        
ErrorHandling:
        CheckMekkoApp
EndSub:
End If

End Sub

'-----------------------------------------------------
'    Send everything to Mekko
'-----------------------------------------------------
Sub SendContentToMekko()
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
                On Error GoTo ErrorHandling
                msg = msg & Addr & delim & cVal & delim & fmt & delim & prefix & delim & frml & Chr(10)
ErrorHandling:
            End If
        Next
        
        SendTextToMekko (msg)
    
    End If
End Sub

'----------------------------------
'     SetRangeValueByMekko
'           (rowBegFmt,     colBegFmt)     - Top left corner of Range with Format
'           (rowCountFmt, colCountFmt) - Dimensions of Range with Format
'           (rowBegVal,     colBegVal)       - Top left corner of Range with content
'           (rowCountVal, colCountVal)   - Dimensions of Range with content
'----------------------------------

Sub SetRangeValueByMekko(ByVal rowBegFmt As String, ByVal colBegFmt As String, _
                                            ByVal rowCountFmt As String, ByVal colCountFmt As String, _
                                            ByVal rowBegVal As String, ByVal colBegVal As String, _
                                            ByVal rowCountVal As String, ByVal colCountVal As String)
    CheckMekkoApp
    If Not IsMekkoActive Or MekkoSenderPort = 0 Then
        SendCmdToMekko ("Error: MekkoSenderPort is not active")
        GoTo EndSub
    End If
    
    Dim t As Single
    t = Timer
    
    Dim rowBegF As Long, rowEndF As Long, rowCountF As Long, lenF As Long
    Dim colBegF As Long, colEndF As Long, colCountF As Long
    
    rowBegF = rowBegFmt
    rowCountF = rowCountFmt
    rowEndF = rowBegF + rowCountF - 1
    colBegF = colBegFmt
    colCountF = colCountFmt
    colEndF = colBegF + colCountF - 1
    lenF = rowCountF * colCountF
    
    Dim rowBegV As Long, rowEndV As Long, rowCountV As Long, lenV As Long
    Dim colBegV As Long, colEndV As Long, colCountV As Long
    
    rowBegV = rowBegVal
    rowCountV = rowCountVal
    rowEndV = rowBegV + rowCountV - 1
    colBegV = colBegVal
    colCountV = colCountVal
    colEndV = colBegV + colCountV - 1
    lenV = rowCountV * colCountV

    Dim cmd As String, data As String
    cmd = "do shell script ""echo '" & Chr(4) & "' | nc 127.0.0.1 " & MekkoSenderPort & """"
    On Error GoTo ErrorHandling
    data = MacScript(cmd)
    
    Dim strArray() As String, delim As String
    delim = Chr(9)
    strArray = Split(data, delim)
    
    If lenV + lenF > (UBound(strArray) - LBound(strArray) + 1) Then
        SendCmdToMekko ("Error: Wrong data lenth")
        GoTo EndSub
    End If
    
    index = LBound(strArray)
    Application.ScreenUpdating = False

    If lenF > 0 Then
        Dim fmtArray() As String
        ReDim fmtArray(1 To rowCountF, 1 To colCountF)
    
        For i = 1 To rowCountF
            For j = 1 To colCountF
                fmtArray(i, j) = strArray(index)
                index = index + 1
            Next j
        Next i
    
        Dim RangeFmt As Range
        Set RangeFmt = Range(Cells(rowBegF, colBegF), Cells(rowEndF, colEndF))
        RangeFmt.NumberFormat = fmtArray
    End If
    
    If lenV > 0 Then
        Dim valArray() As Variant
        ReDim valArray(1 To rowCountV, 1 To colCountV)
        
        For i = 1 To rowCountV
            For j = 1 To colCountV
                valArray(i, j) = strArray(index)
                index = index + 1
            Next j
        Next i
    
        Dim RangeVal As Range
        Set RangeVal = Range(Cells(rowBegV, colBegV), Cells(rowEndV, colEndV))
        RangeVal.Value = valArray
    End If
    
    Application.ScreenUpdating = True
    'SendCmdToMekko ("receieved data in " & Timer - t & " sec")
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
        If wb.Name <> myName Then
            doQuit = False
        Else
            HideByMekko
        End If
    Next
    
    If doQuit Then Application.Quit
End Sub

Private Function IsWorkbookVisible() As Boolean
    IsWorkbookVisible = Workbooks(myName).Windows(1).Visible
End Function

Sub HideByMekko()
    Workbooks(myName).Windows(1).Visible = False
End Sub

Sub UnhideByMekko()
    Workbooks(myName).Windows(1).Visible = True
    'Application.ShowRibbon = True
    
    With ActiveWindow
        .WindowState = xlNormal
        .Height = 360
        .Width = 1500
        .Left = 50
        .Top = 620
        .ScrollColumn = 1
        .ScrollRow = 1
    End With
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
    SendContentToMekko
End Sub
