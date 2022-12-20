Attribute VB_Name = "Module1"
Option Explicit

Const user_password As String = "masterkey"

Private Sub Main()
Set FirebirdDefaultDB = New FirebirdDB
If Not FirebirdDefaultDB.Attach("localhost/3051:" & App.Path & "\DATA.FDB", "SYSDBA", user_password) Then
    MsgBox "Attach Falhou" & vbNewLine & FirebirdDefaultDB.LastErr, vbExclamation
    FirebirdDefaultDB.ClearErr
    If Not FirebirdDefaultDB.Create("localhost/3051:" & App.Path & "\DATA.FDB", "SYSDBA", user_password) Then
        MsgBox "Create Falhou" & vbNewLine & FirebirdDefaultDB.LastErr, vbCritical: Exit Sub
    Else
        MsgBox "Create Funcionou", vbInformation
        FirebirdDefaultDB.Execute "CREATE TABLE TB_TAB (TAB_SEQ INTEGER NOT NULL, TAB_DATA BLOB SUB_TYPE 1 SEGMENT SIZE 80, TAB_WIN1252 VARCHAR(100) CHARACTER SET WIN1252, TAB_UTF8 VARCHAR(100) CHARACTER SET UTF8, TAB_RAW VARCHAR(100), TAB_NUMERIC NUMERIC(15,2), TAB_DECIMAL DECIMAL(15,2), TAB_D DATE, TAB_H TIME, TAB_DH TIMESTAMP, TAB_FLOAT FLOAT, TAB_DOUBLE DOUBLE PRECISION, TB_CHAR CHAR(6) CHARACTER SET ASCII, TAB_VARYING VARCHAR(6) CHARACTER SET ASCII)"
        FirebirdDefaultDB.Execute "ALTER TABLE TB_TAB ADD PRIMARY KEY (TAB_SEQ)"
        FirebirdDefaultDB.Execute "CREATE GENERATOR GE_TAB_SEQ"
        FirebirdDefaultDB.Execute "INSERT INTO TB_TAB (TAB_SEQ, TAB_DATA, TAB_WIN1252, TAB_UTF8, TAB_RAW, TAB_NUMERIC, TAB_DECIMAL, TAB_D, TAB_H, TAB_DH, TAB_FLOAT, TAB_DOUBLE, TB_CHAR, TAB_VARYING) VALUES (100, 'Blob<Data~Blob>Data', 'ã', 'Ãƒ', 'ã', 12.34, 12.34, '26.07.2022', '12:31:32', '26.07.2022 12:31:32', 1/3, 1/6, 'ABC', 'ABCDEF')"
    End If
End If

Dim FB As Firebird
Set FB = New Firebird
Dim B() As Byte
Let B = StrConv("DAMN", vbFromUnicode)

Dim V As Variant
Dim Z As Long

'FB.Prepare "UPDATE OR INSERT INTO TB_TAB (TAB_SEQ,TAB_DATA) VALUES (45,?) MATCHING (TAB_SEQ)"
'FB.Param(0) = B '"DAMN"
'DebugByteArray B
'FB.Execute
'FB.Execute "SELECT * FROM TB_TAB WHERE TAB_SEQ = 45"
'Do While FB.Fetch
'    Debug.Print FB.Row("TAB_SEQ") '; FB.Row("TAB_DATA")
'    DebugString FB.Row("TAB_DATA")
'Loop
'FB.Rollback

FB.Execute "EXECUTE PROCEDURE NEW_PROCEDURE(1)"

Debug.Print FB.RowI(0)

FB.Execute "INSERT INTO TB_TAB (TAB_SEQ) VALUES (GEN_ID(GE_TAB_SEQ,1)) RETURNING (TAB_SEQ)"

Debug.Print FB.RowC

If FB.Fetch Then
    Debug.Print FB.RowI(0)
End If

Debug.Print FirebirdDefaultDB.ExecuteGenId("GE_TAB_SEQ", 0)

FB.Finish

FB.Execute "SELECT 0x180000000 + 69420 POSITIVO,-(0x180000000 + 69420) NEGATIVO FROM RDB$DATABASE"

If FB.Fetch Then
    Debug.Print FB.RowI(0)
    Debug.Print FB.RowI(1)
End If

FB.Finish

FB.Execute "SELECT * FROM TB_TAB WHERE TAB_SEQ = 100"
Do While FB.Fetch
    For Z = 0 To FB.RowC - 1
        V = FB.RowI(Z)
        Debug.Print FB.RowName(Z) & ":"
        If VarType(V) = (vbArray Or vbByte) Then
            DebugByteArray V
        ElseIf VarType(V) = vbString Then
            DebugString V
        Else
            Debug.Print V
        End If
    Next
    Debug.Print FB.Row("TAB_SEQ") '; FB.Row("TAB_DATA")
    DebugString FB.Row("TAB_DATA")
    
Loop
FB.Rollback

FB.Prepare "INSERT INTO TB_TAB (TAB_SEQ,TAB_INT,TAB_RAW) VALUES (GEN_ID(GE_TAB_SEQ,1),?,'SEQUENCE')"

For Z = 1 To 10
    FB.Invoke Z
Next

FB.Execute "SELECT TAB_INT FROM TB_TAB WHERE TAB_RAW = 'SEQUENCE' ORDER BY TAB_INT", , , True

If True Then
    Do While FB.Fetch
        Debug.Print FB.RowI(0)
        If FB.RowI(0) = 7 Then Exit Do
    Loop
    FB.CursorReset
    Do While FB.Fetch
        Debug.Print FB.RowI(0)
    Loop
Else
    FB.CursorSet 0
    Do Until FB.EOF
        Debug.Print FB.RowI(0)
        FB.CursorMove 1
    Loop
    FB.CursorSet 0
    Do Until FB.EOF
        Debug.Print FB.RowI(0)
        FB.CursorMove 1
    Loop
End If

FB.Rollback

FirebirdDefaultDB.Detach
End Sub

Private Function DateOnly(ByVal D As Date) As Date
DateOnly = DateSerial(Year(D), Month(D), Day(D))
End Function

Private Function TimeOnly(ByVal D As Date) As Date
TimeOnly = TimeSerial(Hour(D), Minute(D), Second(D))
End Function

Public Sub DebugByteArray(C As Variant)
Dim Z As Long
Dim S As String
Z = UBound(C)
S = "(" & (Z + 1) & ")["
For Z = 0 To Z
    If Z Then S = S & ", "
    S = S & Chr(C(Z)) & ":" & C(Z)
Next
S = S & "]"
Debug.Print S
End Sub

Public Sub DebugString(C As Variant)
Dim Z As Long
Dim S As String
S = "(" & Len(C) & ")["
For Z = 1 To Len(C)
    If Z > 1 Then S = S & ", "
    S = S & Mid(C, Z, 1) & ":" & AscW(Mid(C, Z, 1))
Next
S = S & "]"
Debug.Print S
End Sub
