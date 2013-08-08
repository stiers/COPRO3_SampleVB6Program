Attribute VB_Name = "ModMain"
' Global Declarations
Public Rs As ADODB.Recordset
Public Con As ADODB.Connection
Public sql As String

Sub Main()

    ' Introducing Global Variables to the Main
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

Public Function Checker(frm As Form) As Boolean

    Dim Control
    Checker = False
    
    For Each Control In frm
    
        If TypeOf Control Is TextBox Or TypeOf Control Is ComboBox Then
        
            If Control.Text = vbNullString Then
            
                If Control.Name = "Text4" Then GoTo jumper
                Checker = True
                MsgBox "All fields required", vbCritical, "Message Alert!"
                Exit Function
                
            End If
        
        End If
    
jumper:

    Next Control

End Function

Public Function mysave(varID, varName, varGen As String, varAge As Integer)

    If Rs.State = 1 Then Rs.Close
    
    sql = "SELECT *FROM TBL_Student WHERE Stud_id='" & varID & "'"
    
    Rs.Open sql, Con
    
    X = MsgBox("Do you want to save this data now?", vbYesNo + vbExclamation, "Message Confirmation")

    If X = vbNo Then Exit Function
    
    If Rs.RecordCount >= 1 Then
        MsgBox "Duplicate of ID", vbCritical
        GoTo jump1
    End If
        
    With Rs
        .AddNew
            .Fields(0) = varID
            .Fields(1) = varName
            .Fields(2) = varGen
            .Fields(3) = varAge
        .Update
    End With
    
    MsgBox "All Data is successfully Saved", vbInformation, "Message Information"
    
jump1:

End Function

Public Function display(grd As MSFlexGrid)

    If Rs.State = 1 Then Rs.Close
    
    sql = "SELECT *FROM tbl_student"
    
    Rs.Open sql, Con
    
    With grd
        .Rows = Rs.RecordCount + 1
        .Cols = Rs.Fields.Count + 1
        
        .TextMatrix(0, 1) = "Student ID": .TextMatrix(0, 2) = "Student Name": .TextMatrix(0, 3) = "Gender"
        .TextMatrix(0, 4) = "Age"
        
        For RowData = 1 To Rs.RecordCount
                        .TextMatrix(RowData, 0) = RowData
            For ColData = 0 To Rs.Fields.Count - 1
            
                .TextMatrix(RowData, ColData + 1) = Rs.Fields(ColData)
            Next ColData
            Rs.MoveNext
        Next RowData

    X = .Width / 4
    .Col = 0
    .ColWidth(0) = 350
    .Col = 2
    .ColWidth(2) = X
    End With

End Function

Public Function m_Path() As String
    
    m_Path = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB1.mdb;Persist Security Info=False"
    
End Function


