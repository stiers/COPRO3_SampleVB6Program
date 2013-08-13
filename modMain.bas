Attribute VB_Name = "modMain"
' Global Variables
Public Rs As ADODB.Recordset
Public Con As ADODB.Connection
Public sql As String

Sub Main()

    Set Rs = New ADODB.Recordset
    Set Con = New ADODB.Connection

    With Rs
        .CursorLocation = adUseClient
        .CursorType = adOpenDynamic
        .LockType = adLockOptimistic
    End With

    Con.Open m_Path

    frmMain.Show
    
End Sub

' Checks all controls if empty
' Prompts user to fill out required fields
Public Function Checker(frm As Form) As Boolean

    Dim Control
    Checker = False
    
    For Each Control In frm
        If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
            If Control.Text = vbNullString Then

                If Control.Name = "Text4" Then GoTo fsChecker
                
                Checker = True
                MsgBox "All fields required", vbCritical, "Message Alert!"
                Exit Function
                
            End If
        End If
    
fsChecker:

    Next Control

End Function

Public Function Add(varStudentID, varStudentName, varStudentGender As String, varStudentAge As Integer)

    If Rs.State = 1 Then Rs.Close
    
    sql = "SELECT * FROM tbl_student WHERE student_id = '" & varID & "'"
    
    Rs.Open sql, Con
    
    x = MsgBox("Do you want to save this data now?", vbYesNo + vbExclamation, "Message Confirmation")

    If x = vbNo Then Exit Function
    
    If Rs.RecordCount >= 1 Then
        MsgBox "Duplicate of ID", vbCritical
        GoTo fsAdd
    End If
        
    With Rs
        .AddNew
            .Fields(0) = varStudentID
            .Fields(1) = varStudentName
            .Fields(2) = varStudentGender
            .Fields(3) = varStudentAge
        .Update
    End With
    
    MsgBox "All Data is successfully Saved", vbInformation, "Message Information"
    
fsAdd:

End Function

Public Function Display(grdMain As MSFlexGrid)

    If Rs.State = 1 Then Rs.Close
    
    sql = "SELECT * FROM tbl_student"
    
    Rs.Open sql, Con
    
    With grdMain
        .Rows = Rs.RecordCount + 1
        .Cols = Rs.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Student ID:"
        .TextMatrix(0, 2) = "Student Name:"
        .TextMatrix(0, 3) = "Gender:"
        .TextMatrix(0, 4) = "Age:"
        
        For RowData = 1 To Rs.RecordCount
            .TextMatrix(RowData, 0) = RowData
            
            For ColData = 0 To Rs.Fields.Count - 1
                .TextMatrix(RowData, ColData + 1) = Rs.Fields(ColData)
            Next ColData

            Rs.MoveNext
        Next RowData

        x = .Width / 4
        
        .Col = 0
        .ColWidth(0) = 350
        
        .Col = 2
        .ColWidth(2) = x
    End With

End Function

Public Function m_Path() As String
    
    m_Path = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\db1.mdb;Persist Security Info=False"
    
End Function



