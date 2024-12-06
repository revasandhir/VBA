Attribute VB_Name = "Module2"
Option Explicit
Sub Main()
    Dim salesRep As String
    Dim state As String
    Dim response As Integer

    ' Loop until the user is tired of making queries.
    Do
        ' For each InputBox, loop until the user enters data.
        Do
            salesRep = InputBox("Enter the last name of a sales rep")
        Loop Until salesRep <> ""
        
        ' Check if the rep is in the list. If not, display an appropriate message and quit.
        Call FindRep(salesRep)
        
        Do
            state = InputBox("Enter a state: Indiana, Ohio, Illinois, Wisconsin, or Michigan")
        Loop Until state <> ""
        
        ' Check if the state has a worksheet. If not, display an appropriate message and quit.
        Call FindState(state)
        
        ' The rep and state must be valid, so find and display the information requested.
        Call FindRepInfo(salesRep, state)
        
        ' Show the SalesRep sheet.
        ThisWorkbook.Sheets("Sales Reps").Activate
        
        ' This is how you can get a Yes/No response from a MsgBox. Yes and No
        ' buttons will appear in the message box, and the value of Response
        ' will be vbYes or vbNo (coded integers).
        response = MsgBox("Do you want to do another search?", vbYesNo)
    Loop Until response = vbNo
End Sub

Sub FindRep(salesRep As String)
    Dim wsSalesReps As Worksheet
    Dim rep As Range

    Set wsSalesReps = ThisWorkbook.Sheets("Sales Reps")

    ' Look for the sales rep in the Sales Reps sheet.
    Set rep = wsSalesReps.Range("B:B").Find(What:=salesRep, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)

    If rep Is Nothing Then
        MsgBox "Sales rep not found!", vbExclamation
        End
    End If
End Sub

Sub FindState(state As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(state)

    If ws Is Nothing Then
        MsgBox "State sheet not found!", vbExclamation
        End
    End If
    On Error GoTo 0
End Sub

Sub FindRepInfo(salesRep As String, state As String)
    Dim ws As Worksheet
    Dim foundRep As Range
    Dim salesTotal As Double
    Dim salesCount As Long
    Dim firstSaleDate As Date
    Dim lastSaleDate As Date
    Dim outputMsg As String

    Set ws = ThisWorkbook.Sheets(state)

    ' Initialize variables
    salesTotal = 0
    salesCount = 0
    firstSaleDate = 0
    lastSaleDate = 0

    ' Look for the sales rep in the state's sheet.
    Set foundRep = ws.Range("B:B").Find(What:=salesRep, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)

    If Not foundRep Is Nothing Then
        Dim firstAddress As String
        firstAddress = foundRep.Address
        
        Do
            salesCount = salesCount + 1
            salesTotal = salesTotal + foundRep.Offset(0, 1).Value
            If firstSaleDate = 0 Then firstSaleDate = foundRep.Offset(0, -1).Value
            lastSaleDate = foundRep.Offset(0, -1).Value
            Set foundRep = ws.Range("B:B").FindNext(foundRep)
        Loop While Not foundRep Is Nothing And foundRep.Address <> firstAddress
        
        outputMsg = salesRep & " made " & salesCount & " sales in " & state & "."
        outputMsg = outputMsg & " The first was on " & Format(firstSaleDate, "mm-dd-yy") & " and the last was on " & Format(lastSaleDate, "mm-dd-yy") & "."
        outputMsg = outputMsg & " The total was for $" & Format(salesTotal, "#,##0.00") & "."
        MsgBox outputMsg, vbInformation
    Else
        MsgBox "Sales rep not found in " & state & "!", vbExclamation
    End If
End Sub

