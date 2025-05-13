Sub UpdateBills()
    ' 🔰 User confirmation
    Dim answer As VbMsgBoxResult
    answer = MsgBox("Are you sure to create new bill?", vbOKCancel + vbQuestion, "Confirm")
    If answer = vbCancel Then Exit Sub ' ❌ Cancel dabane par exit

    Dim ws As Worksheet
    Dim lastDate As Date
    Dim firstDate As Date
    Dim currentMonth As Integer
    Dim nextReqNum As Integer
    Dim cell As Range
    Dim reqText As String
    Dim reqNumber As Integer

    For Each ws In ThisWorkbook.Sheets
        ' Common: Calculate dates
        currentMonth = Month(Now)
        firstDate = DateSerial(Year(Now), currentMonth, 1)
        lastDate = DateSerial(Year(Now), currentMonth + 1, 0)

        ' Get Req# number safely from A1 (🔁 Req# number jaise "Req#3" yahaan likha hota hai)
        reqText = ws.Range("A1").Value      ' 🔹 Ye hai Req# cell — agar change karna ho to "A1" ko replace karo
        If Left(reqText, 4) = "Req#" Then
            On Error Resume Next
            reqNumber = CInt(Mid(reqText, 5))
            On Error GoTo 0
            If reqNumber = 0 Then reqNumber = 3 ' default to Req#3 if parsing fails
            nextReqNum = reqNumber + 1
            ws.Range("A1").Value = "Req#" & nextReqNum     ' 🔹 Yahi par updated Req# likh raha hai — ye bhi "A1" me hi likh raha hai
        End If

        If ws.Name = "Sheet1" Then
            ' Sheet1 specific
            ws.Range("B1").Value = lastDate         ' 🔹 Sheet1 me sirf current month ka last date B1 me set ho raha hai
            ws.Range("D1").Value = ws.Range("C1").Value   ' 🔹 C1 ka value D1 me copy ho raha hai sirf Sheet1 me
        Else
            ' Other sheets
            ws.Range("A2").Value = firstDate        ' 🔹 A2 me current month ki pehli date ja rahi hai (jaise 5/1/2025)
            ws.Range("A3").Value = lastDate         ' 🔹 A3 me current month ki last date ja rahi hai (jaise 5/31/2025)

            ' Copy values from C1:C30 to B1:B30
            For Each cell In ws.Range("C1:C30")     ' 🔹 Ye column C (C1 se C30) me se value le raha hai
                cell.Offset(0, -1).Value = cell.Value    ' 🔹 Us value ko uske left wale cell me (Column B) paste kar raha hai
            Next cell

            ' Set D1:D30 to 0
            ws.Range("D1:D30").Value = 0            ' 🔹 Column D (D1 to D30) me 0 daal raha hai
        End If
    Next ws
End Sub

