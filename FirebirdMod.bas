Attribute VB_Name = "FirebirdMod"
Option Explicit

'Config:

Public Const CommitByDefault As Boolean = True
Public Const CacheBlob As Boolean = True
Public Const StoreRowsByDefault As Boolean = True
Public Const CoalesceByDefault As Boolean = True

'Instancia Padrão do FirebirdDB, para novas instancia do Firebird
Public FirebirdDefaultDB As FirebirdDB

'Publico,mas reservado para Firebird, FirebirdDB e FirebirdMod
Public status_vector(255) As Long

Private Declare Function isc_interprete Lib "fbclient.dll" (ByVal Buffer As Long, status_vector_pointer As Long) As Long
Private Declare Sub memcls Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Dim OpenDBs() As FirebirdDB
Dim OpenDBc As Long

'0 = Firebird Error
'1 = Database Not Attached / Empty Firebird Class / tr_handle = 0
'2 = Arg Index Out Of Bounds (Param)
'3 = Field Name Not Found    (Row)
'4 = Variant Import Error    (Param)
'5 = Variant Export Error    (Row)
Function z_FBEC(Optional ByVal Z As Long) As String
Dim X As Long
'Essa funcao captura todos os erros que ocorrem com Firebird e FirebirdMod
If Z = 0 Then
    If Not (status_vector(0) = 1 And status_vector(1) <> 0) Then Exit Function
    Dim Buffer As String
    Z = VarPtr(status_vector(0))
    Buffer = String$(1024, 0)
    Do
        X = isc_interprete(StrPtr(Buffer), Z)
        If X = 0 Then Exit Do
        z_FBEC = z_FBEC & Left$(StrConv(Buffer, vbUnicode), X) & vbNewLine
        memcls ByVal StrPtr(Buffer), LenB(Buffer)
    Loop
    memcls status_vector(0), LenB(status_vector(0)) * (UBound(status_vector) + 1)
Else
    z_FBEC = "Firebird Wrapper Error" & vbNewLine
    Select Case Z
    Case 1: z_FBEC = z_FBEC & "Handle Nulo"
    Case 2: z_FBEC = z_FBEC & "Nenhum FirebirdDB Associado"
    Case 3: z_FBEC = z_FBEC & "Banco Não Aberto"
    Case 4: z_FBEC = z_FBEC & "Sql Não Criada"
    Case 5: z_FBEC = z_FBEC & "Index De Argumento Invalido"
    Case 6: z_FBEC = z_FBEC & "Campo Não Econtrado"
    Case 7: z_FBEC = z_FBEC & "Não Foi Possivel Transformar a Variante Em Um Parametro Sql"
    Case 8: z_FBEC = z_FBEC & "Não Foi Possivel Transformar o Valor Sql Em Uma Variante"
    Case 9: z_FBEC = z_FBEC & "Não é Possível Retroceder o Cursor Sem Que Execute Tenha Sido Chamado Com StoreRows = True"
    Case 10: z_FBEC = z_FBEC & "Row Foi Chamado Com EOF = True Ou Sem Fetch Ser Chamado"
    End Select
End If
'Debug.Assert False
Debug.Print z_FBEC
End Function

Sub z_FBAD(FirebirdDB As FirebirdDB)
Dim Z As Long
If FirebirdDB.Handle Then
    OpenDBc = OpenDBc + 1
    ReDim Preserve OpenDBs(OpenDBc - 1)
    Set OpenDBs(OpenDBc - 1) = FirebirdDB
ElseIf OpenDBc = 1 Then
    If OpenDBs(0) Is FirebirdDB Then
        OpenDBc = 0
        Erase OpenDBs
    End If
Else
    For Z = 0 To OpenDBc - 1
        If OpenDBs(Z) Is FirebirdDB Then
            Set OpenDBs(Z) = OpenDBs(OpenDBc - 1)
            OpenDBc = OpenDBc - 1
            ReDim Preserve OpenDBs(OpenDBc - 1)
        End If
    Next
End If
End Sub

Public Sub FirebirdDetachAllDatabases()
Dim Z As Long
Dim LastDetached As Long
Do While Z < OpenDBc
    LastDetached = ObjPtr(OpenDBs(Z))
    OpenDBs(Z).Detach
    If LastDetached = ObjPtr(OpenDBs(Z)) Then Z = Z + 1
Loop
End Sub
