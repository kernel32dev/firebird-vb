VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Firebird"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub memcls Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Const METADATALENGTH As Long = 30
Const DIALECT As Integer = 3

Private Type XSQLDA
    version As Integer 'SQLDA_CURRENT_VERSION = 1
    sqldaid(8) As Byte
    sqldabc As Long
    sqln As Integer 'Capacity
    sqld As Integer 'Size / Needed Count 'bytes 18 and 19
End Type

Private Type XSQLVAR
    sqltype As Integer 'SqlTypeEnum
    sqlscale As Integer
    'sqlprecision As Integer 'Ou sqlsubtype ou sqlscale ou sqlprecision não existe para essa versão do firebird
    sqlsubtype As Integer
    sqllen As Integer
    sqldata As Long 'Pointer To Char
    sqlind As Long 'Pointer To Short
    sqlname_len As Integer
    sqlname(METADATALENGTH) As Byte
    relname_len As Integer
    relname(METADATALENGTH) As Byte
    ownname_len As Integer
    ownname(METADATALENGTH) As Byte
    aliasname_len As Integer
    aliasname(METADATALENGTH) As Byte
End Type

Private IndNull As Integer ' = -1
Private IndNotNull As Integer ' = 0

Public Enum StatementKindEnum
st_kind_none = 0
st_kind_select = 1
st_kind_insert = 2
st_kind_update = 3
st_kind_delete = 4
st_kind_ddl = 5
st_kind_get_segment = 6
st_kind_put_segment = 7
st_kind_exec_procedure = 8
st_kind_start_trans = 9
st_kind_commit = 10
st_kind_rollback = 11
st_kind_select_for_upd = 12
st_kind_set_generator = 13
st_kind_savepoint = 14
End Enum

Private Enum SqlTypeEnum
SQL_TEXT = 452
SQL_VARYING = 448
SQL_SHORT = 500
SQL_LONG = 496
SQL_FLOAT = 482
SQL_DOUBLE = 480
SQL_D_FLOAT = 530
SQL_TIMESTAMP = 510
SQL_BLOB = 520
SQL_ARRAY = 540
SQL_QUAD = 550
SQL_TYPE_TIME = 560
SQL_TYPE_DATE = 570
SQL_INT64 = 580
SQL_BOOL = 32764
End Enum

'Essa Classe Representa Uma Query, Que, Esta Associada A Uma Transação (TransactionHandle)
'Ha Apenas Uma Query Por Classe, Mas Ela Pode Compartilhar a Transação
'Cada Query Pode Ser Executada Varias Vezes Mudando Parametros

'A Connecção Com o Banco de dados é global

'Private Declare Function isc_start_multiple Lib "fbclient.dll" (status_vector As Long, tr_handle As Long, ByVal amount_of_tebs As Integer, tebs As IscTeb) As Long
'Private Declare Function isc_start_transaction CDecl Lib "fbclient.dll" (status_vector As Long, tr_handle As Long, ByVal amount_of_tebs As Integer, db_handle As Long, ByVal tpb_length As Long, ByVal tpb As String) As Long
'Private Declare Function isc_commit_transaction Lib "fbclient.dll" (status_vector As Long, tr_handle As Long) As Long
'Private Declare Function isc_rollback_transaction Lib "fbclient.dll" (status_vector As Long, tr_handle As Long) As Long


Private Declare Function isc_dsql_allocate_statement Lib "fbclient.dll" (status_vector As Long, db_handle As Long, statement_handle As Long) As Long
Private Declare Function isc_dsql_free_statement Lib "fbclient.dll" (status_vector As Long, st_handle As Long, Optional ByVal stFreeMode As Integer = 2) As Long
'stFreeMode = 1 = Drop Cursor used To Fetch The Data, Making The Statment Available Again
'stFreeMode = 2 = Drop Statement itself, For Every isc_dsql_allocate_statement, One isc_dsql_free_statement(2) is needed

Private Declare Function isc_dsql_sql_info Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 st_handle As Long, _
                                                 ByVal req_length As Integer, _
                                                 ByVal req As Long, _
                                                 ByVal res_length As Integer, _
                                                 ByVal res As Long) As Long

Private Declare Function isc_dsql_prepare Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 tr_handle As Long, _
                                                 st_handle As Long, _
                                                 ByVal sql_length As Integer, _
                                                 ByVal SQL As String, _
                                                 ByVal DIALECT As Integer, _
                                                 outXSQLDA As Any) As Long

Private Declare Function isc_dsql_describe_bind Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 st_handle As Long, _
                                                 ByVal da_version As Integer, _
                                                 inXSQLDA As Any) As Long

Private Declare Function isc_dsql_describe Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 st_handle As Long, _
                                                 ByVal da_version As Integer, _
                                                 outXSQLDA As Any) As Long

Private Declare Function isc_dsql_execute Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 tr_handle As Long, _
                                                 st_handle As Long, _
                                                 ByVal da_version As Integer, _
                                                 inXSQLDA As Any) As Long

Private Declare Function isc_dsql_execute2 Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 tr_handle As Long, _
                                                 st_handle As Long, _
                                                 ByVal da_version As Integer, _
                                                 inXSQLDA As Any, _
                                                 outXSQLDA As Any) As Long

Private Declare Function isc_dsql_set_cursor_name Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 st_handle As Long, _
                                                 ByVal cursor_name As String, _
                                                 outXSQLDA As Any) As Long

'da_version = 1
'dialect = 3

Private Declare Function isc_dsql_fetch Lib "fbclient.dll" (status_vector As Long, st_handle As Long, ByVal da_version As Long, outXSQLDA As Any) As Long

'BLOB READING
Private Declare Function isc_open_blob2 Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 db_handle As Long, _
                                                 tr_handle As Long, _
                                                 blob_handle As Long, _
                                                 ISC_QUAD As Any, _
                                                 ByVal BPBLen As Integer, _
                                                 BPB As Any) As Long

Private Declare Function isc_create_blob2 Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 db_handle As Long, _
                                                 tr_handle As Long, _
                                                 blob_handle As Long, _
                                                 ISC_QUAD As Any, _
                                                 ByVal BPBLen As Integer, _
                                                 BPB As Any) As Long

Private Declare Function isc_get_segment Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 blob_handle As Long, _
                                                 bytesRead As Integer, _
                                                 ByVal segmentLength As Integer, _
                                                 segment As Any) As Long

Private Declare Function isc_put_segment Lib "fbclient.dll" ( _
                                                 status_vector As Long, _
                                                 blob_handle As Long, _
                                                 ByVal segmentLength As Integer, _
                                                 segment As Any) As Long

Private Declare Function isc_close_blob Lib "fbclient.dll" (status_vector As Long, blob_handle As Long) As Long
Private Declare Function isc_cancel_blob Lib "fbclient.dll" (status_vector As Long, blob_handle As Long) As Long

Dim db_instance As FirebirdDB

Dim tr_handle As Long 'Compartilhado
Dim st_handle As Long 'Privado

Dim StatementText As String
Dim StatementKind As StatementKindEnum

Dim XSQLIN() As Byte
Dim XSQLOUT() As Byte
Dim Fetching As Boolean

Dim Src() As Variant
Dim SrcBuff() As String
Dim MaxSrc As Long

Dim Dest() As Variant
Dim Fields() As String
Dim IndArr() As Integer
Dim MaxDest As Long

Dim DestBuff As String

Dim Storing As Boolean
Dim StoreBuff() As Long
Dim StoreItemBytes As Long 'Size In Bytes
Dim StoreItemLongs As Long 'Size In Longs
Dim StoreRowC As Long 'Count
Dim StoreRowCC As Long 'Capacity

Dim CursorPosition As Long

Private Type BlobChacheSlot
Quad(1) As Long
B() As Byte
Max As Long
End Type

Dim Cache() As BlobChacheSlot
Dim MaxC As Long

Public LastErr As String

Property Get Err() As Boolean
Err = LenB(LastErr)
End Property

Public Sub ClearErr()
LastErr = vbNullString
End Sub

Private Sub SetErr(Optional ByVal Z As Long)
LastErr = z_FBEC(Z)
End Sub

Property Get Statement() As String
Statement = StatementText
End Property

Property Get StatementType() As StatementKindEnum
StatementType = StatementKind
End Property

Private Sub ResetST()
'Does Not Change db_instance or tr_handle
If st_handle Then
isc_dsql_free_statement status_vector(0), st_handle
StatementText = vbNullString
st_handle = 0
End If
Erase Src, SrcBuff, Dest, Fields, IndArr, XSQLIN, XSQLOUT
MaxSrc = -1:  MaxDest = -1
DestBuff = vbNullString
StatementText = vbNullString
StatementKind = 0
Fetching = False
Storing = False
StoreRowC = 0
StoreRowCC = 0
Erase StoreBuff
StoreItemBytes = 0
StoreItemLongs = 0
MaxC = -1
Erase Cache
CursorPosition = -1
End Sub

Public Function Setup(ByVal Transaction As Long, Database As FirebirdDB) As Boolean
Dim db_temp As FirebirdDB
Dim tr_temp As Long

ResetST

If Transaction < 0 Then
    tr_handle = 0
    Set db_instance = Database
    Setup = True
    Exit Function
End If

If Not Database Is Nothing Then
    Set db_temp = Database 'Use Parameter
ElseIf Not db_instance Is Nothing Then
    Set db_temp = db_instance 'Use Previous
ElseIf Not FirebirdDefaultDB Is Nothing Then
    Set db_temp = FirebirdDefaultDB 'Use Default
Else
    If tr_handle <> 0 And Not db_instance Is Nothing Then
        db_instance.z_FBTA tr_handle, ObjPtr(Me), False
    End If
    tr_handle = 0
    Set db_instance = Nothing
    SetErr 2: Exit Function 'Error
End If

If db_temp.Handle = 0 Then
    If tr_handle <> 0 And Not db_instance Is Nothing Then
        db_instance.z_FBTA tr_handle, ObjPtr(Me), False
    End If
    tr_handle = 0
    Set db_instance = Nothing
    SetErr 3: Exit Function
End If

If Transaction Then
    tr_temp = Transaction 'User Parameter
ElseIf Database Is Nothing And db_temp Is db_instance Then
    tr_temp = tr_handle 'Use Previous
Else
    tr_temp = db_temp.Start() 'Use Default (create)
    If tr_temp = 0 Then
        If tr_handle <> 0 And Not db_instance Is Nothing Then
            db_instance.z_FBTA tr_handle, ObjPtr(Me), False
        End If
        tr_handle = 0
        Exit Function
    End If
End If

If tr_handle <> tr_temp Then
    If Not db_instance Is Nothing Then db_instance.z_FBTA tr_handle, ObjPtr(Me), False
    db_temp.z_FBTA tr_temp, ObjPtr(Me), True
End If

tr_handle = tr_temp
Set db_instance = db_temp

Setup = True

End Function

'Retorna False Se Tiver Ocorrido Um Erro
Public Function Prepare(SqlStatement As String, Optional ByVal Transaction As Long, Optional Database As FirebirdDB) As Boolean
Dim XSQLDA As XSQLDA
Dim Var As XSQLVAR
Dim Z As Long, X As Long
Dim DoCursor As Boolean

If Not Setup(Transaction, Database) Then
    Exit Function
End If

If LenB(SqlStatement) = 0 Then
    Prepare = True
    Exit Function
End If

If isc_dsql_allocate_statement(status_vector(0), db_instance.Handle, st_handle) Then
    st_handle = 0
    GoTo free
    Exit Function
End If

XSQLDA.version = 1
If isc_dsql_prepare(status_vector(0), tr_handle, st_handle, Len(SqlStatement), SqlStatement, DIALECT, XSQLDA) Then GoTo free

Const isc_info_sql_stmt_type As Byte = 21

Dim SqlInfo(8) As Byte

If isc_dsql_sql_info(status_vector(0), st_handle, 1, VarPtr(isc_info_sql_stmt_type), 8, VarPtr(SqlInfo(0))) Then GoTo free

If SqlInfo(0) = isc_info_sql_stmt_type Then
    Debug.Print SqlInfo(1)
    Debug.Print SqlInfo(2)
    StatementKind = SqlInfo(3)
    If StatementKind < 1 Or StatementKind > 14 Then
        StatementKind = 0
    End If
    'StatementKind = 1
End If

XSQLDA.sqln = XSQLDA.sqld
ReDim XSQLOUT(LenB(XSQLDA) + LenB(Var) * XSQLDA.sqld)

memcpy XSQLOUT(0), XSQLDA, LenB(XSQLDA)

If XSQLDA.sqld <= 0 Then
    MaxDest = -1
    Erase Dest
Else
    If isc_dsql_describe(status_vector(0), st_handle, 1, XSQLOUT(0)) Then GoTo free
    MaxDest = XSQLDA.sqld - 1
    ReDim Dest(MaxDest)
    ReDim Fields(MaxDest)
    ReDim IndArr(MaxDest)
    For Z = 0 To MaxDest
        memcpy Var, ByVal VarPtr(XSQLOUT(0)) + LenB(XSQLDA) + LenB(Var) * Z, LenB(Var)
        Fields(Z) = String$(Var.sqlname_len, 0)
        'Var.sqltype = Var.sqltype And Not 1
        Var.sqlind = VarPtr(IndArr(Z))
        memcpy ByVal StrPtr(Fields(Z)), Var.sqlname(0), Var.sqlname_len
        LSet Fields(Z) = StrConv(Fields(Z), vbUnicode)
        If (Var.sqltype And Not 1&) = SQL_VARYING Then
            X = X + (Var.sqllen + 3 + 2) And Not 3& '4 Byte Align (+ 2 bytes)
        Else
            X = X + (Var.sqllen + 3) And Not 3& '4 Byte Align
        End If
        memcpy ByVal VarPtr(XSQLOUT(0)) + LenB(XSQLDA) + LenB(Var) * Z, Var, LenB(Var)
    Next
    'X Is Total Bytes Required
    'X Is 4 Byte Aligned
    DestBuff = String$(X \ 2, 0)
    
    'Buff Item Size is LenB(DestBuff) + (MaxDest + 1) * 2) + (optional 2 Bytes)
    'First LenB(DestBuff) Bytes Is The DestBuffer
    'The Next (MaxDest + 1) * 2 Bytes Is The IndArr
    'If MaxDest + 1 Is Odd, or MaxDest Is Even, 2 unused extra bytes should be added, for alignment
    'The Final Size must be 4 Byte Aligned
    Debug.Assert (CLng(LenB(DestBuff)) And 3) = 0
    StoreItemBytes = LenB(DestBuff) + (MaxDest + 1) * 2
    If (MaxDest And 1) = 0 Then StoreItemBytes = StoreItemBytes + 2
    Debug.Assert (CLng(StoreItemBytes) And 3) = 0
    StoreItemLongs = StoreItemBytes \ 4
    
    X = StrPtr(DestBuff)
    For Z = 0 To MaxDest
        memcpy Var, ByVal VarPtr(XSQLOUT(0)) + LenB(XSQLDA) + LenB(Var) * Z, LenB(Var)
        Var.sqldata = X
        X = X + (Var.sqllen + 3) And Not 3 '4 Byte Align
        memcpy ByVal VarPtr(XSQLOUT(0)) + LenB(XSQLDA) + LenB(Var) * Z, Var, LenB(Var)
    Next
End If

memcls XSQLDA, LenB(XSQLDA)

XSQLDA.version = 1
If isc_dsql_describe_bind(status_vector(0), st_handle, 1, XSQLDA) Then GoTo free
XSQLDA.sqln = XSQLDA.sqld
ReDim XSQLIN(LenB(XSQLDA) + LenB(Var) * XSQLDA.sqld)
memcpy XSQLIN(0), XSQLDA, LenB(XSQLDA)
If XSQLDA.sqld <= 0 Then
    MaxSrc = -1
    Erase Src
Else
    If isc_dsql_describe_bind(status_vector(0), st_handle, 1, XSQLIN(0)) Then GoTo free
    MaxSrc = XSQLDA.sqld - 1
    ReDim Src(MaxSrc)
    ReDim SrcBuff(MaxSrc)
    For Z = 0 To MaxSrc
        memcpy Var, ByVal VarPtr(XSQLIN(0)) + LenB(XSQLDA) + LenB(Var) * Z, LenB(Var)
        Var.sqlind = VarPtr(IndNull)
        memcpy ByVal VarPtr(XSQLIN(0)) + LenB(XSQLDA) + LenB(Var) * Z, Var, LenB(Var)
        Src(Z) = Null
    Next
End If

StatementText = SqlStatement
Prepare = True
Exit Function
free: SetErr
'dont free transaction
'If tr_handle Then db_instance.z_FBTA tr_handle, ObjPtr(Me), False
'tr_handle = 0: Set db_instance = Nothing
ResetST
End Function

'Executa a query uma vez, com os parametros especificados
'Se houver dados que não foram lidos da ultima vez, são descartados
'Se o comando executado retornar dados, eles poderão ser obtidos
'Se SqlStatement For Especificado, Prepare e Chamado Antes de Ser Executado
'Retorna False Se Tiver Ocorrido Um Erro
Public Function Execute(Optional SqlStatement As String, Optional ByVal Transaction As Long, Optional Database As FirebirdDB, Optional ByVal StoreRows As Boolean = StoreRowsByDefault) As Boolean
If StrPtr(SqlStatement) Then
    If Not Prepare(SqlStatement, Transaction, Database) Then Exit Function
End If
If st_handle = 0 Then SetErr 4: Exit Function
If Fetching Then
    isc_dsql_free_statement status_vector(0), st_handle, 1
    Fetching = False
End If
If StatementKind = st_kind_exec_procedure And MaxDest <> -1 Then
    If isc_dsql_execute2(status_vector(0), tr_handle, st_handle, 1, XSQLIN(0), XSQLOUT(0)) Then SetErr: Exit Function
Else
    If isc_dsql_execute(status_vector(0), tr_handle, st_handle, 1, XSQLIN(0)) Then SetErr: Exit Function
End If
If MaxDest <> -1 Then
    If StatementKind = st_kind_select Then
        Fetching = 0 = isc_dsql_set_cursor_name(status_vector(0), st_handle, "dyn_cursor_" & ObjPtr(Me), XSQLIN(0))
        If Fetching Then
            Storing = StoreRows
        Else
            SetErr
        End If
    Else
        Fetching = StatementKind = st_kind_exec_procedure
    End If
End If
Execute = True
End Function

Public Function Invoke(ParamArray Args() As Variant) As Boolean
Dim Z As Long
If UBound(Args) <= MaxSrc Then
    For Z = 0 To UBound(Args)
        Param(Z) = Args(Z)
    Next
    For Z = UBound(Args) + 1 To MaxSrc
        Param(Z) = Null
    Next
    Invoke = Execute
Else
    SetErr 5
    Exit Function
End If
End Function

Public Function Fetch() As Boolean
Dim Z As Long
Dim SIS As Long
If st_handle = 0 Then SetErr 4: Exit Function
If Not Fetching Then
    If Storing And CursorPosition + 1 < StoreRowC Then
        CursorSet CursorPosition + 1
        Fetch = CursorPosition < StoreRowC
    End If
    Exit Function
End If
If StatementKind <> st_kind_select Then
    CursorPosition = 0
    Fetch = True
    Fetching = False
    Exit Function
End If
If Not Storing Then
    'erase blob cache, as row is now unreachable
    MaxC = -1
    Erase Cache
End If
If CursorPosition + 1 < StoreRowC Then
    CursorSet CursorPosition + 1
    Fetch = True
    Exit Function
End If
Z = isc_dsql_fetch(status_vector(0), st_handle, 1, XSQLOUT(0))
If Z = 100 Then
    Fetching = False
ElseIf Z Then
    Fetching = False
    SetErr
Else
    Fetch = True
    CursorPosition = CursorPosition + 1
    If Storing Then
        'store current DestBuffer
        If StoreRowCC = StoreRowC Then
            'Increase Capacity
            If StoreRowCC Then
                StoreRowCC = StoreRowCC * 2
            Else
                StoreRowCC = 16
            End If
            ReDim Preserve StoreBuff(StoreItemLongs * StoreRowCC)
        End If
        memcpy StoreBuff(StoreItemLongs * StoreRowC), ByVal StrPtr(DestBuff), LenB(DestBuff)
        memcpy StoreBuff(StoreItemLongs * StoreRowC + LenB(DestBuff) \ 4), IndArr(0), (MaxDest + 1) * 2
        StoreRowC = StoreRowC + 1
    End If
End If
End Function

Property Get Cursor() As Long
Cursor = CursorPosition
End Property

Property Let Cursor(ByVal Index As Long)
If Index >= 0 Then
    CursorSet Index
Else
    CursorSet 0
End If
End Property

Sub CursorMove(Optional ByVal Offset As Long = 1)
If Offset + CursorPosition >= 0& Then
    CursorSet Offset + CursorPosition
Else
    CursorSet 0&
End If
End Sub

Sub CursorSet(Optional ByVal Index As Long = 0)
If Storing Then
    If Index < 0& Then
        If Fetching Then
            Do While Fetch 'Fetch All
            Loop
        End If
        Index = Index + StoreRowC
        If Index < 0& Then Index = StoreRowC
    End If
    If CursorPosition = Index Then Exit Sub
    If Index >= StoreRowC Then
        If Fetching Then
            Do While Fetch
                If Index < StoreRowC Then Exit Do
            Loop
        End If
        If Index >= StoreRowC Then
            CursorPosition = StoreRowC
        End If
    End If
    memcpy ByVal StrPtr(DestBuff), StoreBuff(StoreItemLongs * Index), LenB(DestBuff)
    memcpy IndArr(0), StoreBuff(StoreItemLongs * Index + LenB(DestBuff) \ 4), (MaxDest + 1) * 2
    CursorPosition = Index
Else
    If Not Fetching Then
        'Não ha mais nada para pegar, mas tentamos voltar para algum registro anterior a EOF
        If Index >= 0 And Index < CursorPosition Then SetErr 9
        Exit Sub
    End If
    If Index < 0& Or Index < CursorPosition Then
        SetErr 9
        Exit Sub
    End If
    If CursorPosition = Index Then Exit Sub
    If Fetching Then
        Do While Fetch
            If Index >= CursorPosition Then Exit Do
        Loop
    End If
End If
End Sub

Sub CursorReset()
If MaxSrc > 0 And StatementKind = st_kind_exec_procedure Then
    Fetching = True
End If
CursorPosition = -1
End Sub

Property Get EOF() As Boolean
If Fetching Then
    EOF = False
ElseIf Storing Then
    '...
    EOF = CursorPosition = StoreRowC
Else
    EOF = True
End If
End Property

'Da um Commit na transação para esta classe e com todas as quais compartilha sua transação
Public Function Commit() As Boolean
If tr_handle Then Commit = db_instance.Commit(tr_handle)
End Function

'Da um Rollback na transação para esta classe e com todas as quais compartilha sua transação
'Se não hover Transação faz nada e retorna false
Public Function Rollback() As Boolean
If tr_handle Then Rollback = db_instance.Rollback(tr_handle)
End Function

'Da um Commit ou Rollback dependendo da constante CommitByDefault
Public Function Forget() As Boolean
If tr_handle Then Forget = db_instance.Forget(tr_handle)
End Function

'Da um Commit ou Rollback dependendo se houve erros
'Lembrando que pode ter tido erros em outras instancias
Public Function Finish() As Boolean
If tr_handle Then
    If Not Err Then
        Finish = db_instance.Commit(tr_handle)
    Else
        Finish = db_instance.Rollback(tr_handle)
    End If
End If
End Function

Property Get Database() As FirebirdDB
Set Database = db_instance
End Property

Property Get Transaction() As Long
Transaction = tr_handle
End Property

Function SetParams(ParamArray Args() As Variant) As Boolean
Dim Z As Long
If UBound(Args) <= MaxSrc Then
    For Z = 0 To UBound(Args)
        Param(Z) = Args(Z)
    Next
    For Z = UBound(Args) + 1 To MaxSrc
        Param(Z) = Null
    Next
    SetParams = True
Else
    SetErr 5
End If
End Function

Property Get ParamC() As Long
ParamC = MaxSrc + 1
End Property

Property Get Param(ByVal Index As Long) As Variant
If Index <= MaxSrc And Index >= 0 Then Param = Src(Index) Else SetErr 5
End Property

Property Let Param(ByVal Index As Long, Val As Variant)
Dim Tmp As XSQLDA
Dim Var As XSQLVAR
Dim Z As Long
If Index > MaxSrc Or Index < 0 Then SetErr 5: Exit Property
Src(Index) = Val

memcpy Var, ByVal VarPtr(XSQLIN(0)) + LenB(Tmp) + LenB(Var) * Index, LenB(Var)
SrcBuff(Index) = vbNullString
If VarType(Val) = vbNull Or VarType(Val) = vbEmpty Then
    Var.sqldata = 0
    Var.sqlind = VarPtr(IndNull)
ElseIf Pack(Src(Index), SrcBuff(Index), Var) Then
    Var.sqldata = StrPtr(SrcBuff(Index))
    Var.sqlind = VarPtr(IndNotNull)
Else
    SrcBuff(Index) = vbNullString
    Var.sqlind = VarPtr(IndNull)
    SetErr 7
End If
memcpy ByVal VarPtr(XSQLIN(0)) + LenB(Tmp) + LenB(Var) * Index, Var, LenB(Var)
End Property

Property Get RowC() As Long
RowC = MaxDest + 1
End Property

Property Get Row(Field As String, Optional ByVal Coalesce As Boolean = CoalesceByDefault) As Variant
Dim Z As Long
For Z = 0 To MaxDest
    If Fields(Z) Like Field Then
        Row = RowI(Z)
        Exit Property
    End If
Next
SetErr 6
End Property

Property Get RowName(ByVal Index As Long) As String
If Index >= 0 And Index <= MaxDest Then RowName = Fields(Index)
End Property

Property Get RowI(ByVal Index As Long, Optional ByVal Coalesce As Boolean = CoalesceByDefault) As Variant
Dim Z As Long
Dim Tmp As XSQLDA
Dim Var As XSQLVAR
If Index < 0 Or Index > MaxDest Then
    SetErr 6
ElseIf EOF Or CursorPosition = -1 Then
    SetErr 10
Else
    memcpy Var, ByVal VarPtr(XSQLOUT(0)) + LenB(Tmp) + LenB(Var) * Index, LenB(Var)
    RowI = Unpack(Var, Coalesce)
End If
End Property

Private Sub Class_Initialize()
MaxSrc = -1
MaxDest = -1
IndNull = -1
MaxC = -1
End Sub

Private Sub Class_Terminate()
If st_handle Then isc_dsql_free_statement status_vector(0), st_handle, 2
If tr_handle Then db_instance.z_FBTA tr_handle, ObjPtr(Me), False
End Sub

Private Function Pack(V As Variant, Buff As String, Var As XSQLVAR) As Boolean
'Copy V's Value Into Buff, To Be In A Readble Format For Firebird
'If Buff Is Set To vbNullString, Then It Is Null
'Returns False On Failure
Pack = True
Select Case VarType(V)
'Case vbInteger, vbByte 'Estava causando um problema
'    Var.sqltype = SQL_SHORT: Buff = " "
'    memcpy ByVal StrPtr(Buff), CInt(V), 2
Case vbByte, vbInteger, vbLong
    Var.sqltype = SQL_LONG: Buff = "  "
    memcpy ByVal StrPtr(Buff), CLng(V), 4
Case vbSingle
    Var.sqltype = SQL_FLOAT: Buff = "  "
    memcpy ByVal StrPtr(Buff), CSng(V), 4
Case vbDouble
    Var.sqltype = SQL_DOUBLE: Buff = "    "
    memcpy ByVal StrPtr(Buff), CDbl(V), 8
Case vbString
    Var.sqltype = SQL_VARYING
    Buff = StrConv(V, vbFromUnicode)
    'Buff = " " & StrConv(V, vbFromUnicode)
    'memcpy ByVal StrPtr(Buff), CInt(Len(V)), 2
Case vbDate
    If Hour(V) = 0 And Minute(V) = 0 And Second(V) = 0 Then
        Var.sqltype = SQL_TYPE_DATE 'Date Only
        Buff = "  "
        memcpy ByVal StrPtr(Buff), EncodeDate(V), 4
    ElseIf Year(V) = 1899 And Month(V) = 12 And Day(V) = 30 Then
        Var.sqltype = SQL_TYPE_TIME 'Time Only
        Buff = "  "
        memcpy ByVal StrPtr(Buff), EncodeTime(V), 4
    Else
        Var.sqltype = SQL_TIMESTAMP 'Date And Time
        Buff = "    "
        memcpy ByVal StrPtr(Buff), EncodeDate(V), 4
        memcpy ByVal StrPtr(Buff) + 4, EncodeTime(V), 4
    End If
Case vbArray + vbByte
    'BLOB
    Var.sqltype = SQL_BLOB
    Buff = "    "
    Pack = WriteBlob(V, StrPtr(Buff))
Case Else
    Pack = False
End Select
End Function

Private Sub ApplyScale(Value As Variant, ByVal DecimalScale As Integer)
Select Case DecimalScale
Case 0: Exit Sub
Case -1: Value = Value * 0.1@
Case -2: Value = Value * 0.01@
Case -3: Value = Value * 0.001@
Case -4: Value = Value * 0.0001@
Case -5: Value = Value * 0.00001
Case -6: Value = Value * 0.000001
Case -7: Value = Value * 0.0000001
Case -8: Value = Value * 0.00000001
Case -9: Value = Value * 0.000000001
Case -10: Value = Value * 0.0000000001
Case -11: Value = Value * 0.00000000001
Case -12: Value = Value * 0.000000000001
Case -13: Value = Value * 0.0000000000001
Case -14: Value = Value * 0.00000000000001
Case -15: Value = Value * 0.000000000000001
Case -16: Value = Value * 1E-16
Case -17: Value = Value * 1E-17
Case -18: Value = Value * 1E-18
Case -19: Value = Value * 1E-19
Case -20: Value = Value * 1E-20
Case -21: Value = Value * 1E-21
Case -22: Value = Value * 1E-22
Case -23: Value = Value * 1E-23
Case -24: Value = Value * 1E-24
Case -25: Value = Value * 1E-25
Case 1: Value = Value * 10@
Case 2: Value = Value * 100@
Case 3: Value = Value * 1000@
Case 4: Value = Value * 10000@
Case 5: Value = Value * 100000@
Case 6: Value = Value * 1000000@
Case 7: Value = Value * 10000000@
Case 8: Value = Value * 100000000@
Case 9: Value = Value * 1000000000@
Case 10: Value = Value * 10000000000@
Case 11: Value = Value * 100000000000@
Case 12: Value = Value * 1000000000000@
Case 13: Value = Value * 10000000000000@
Case 14: Value = Value * 100000000000000@
Case 15: Value = Value * 1E+15
Case 16: Value = Value * 1E+16
Case 17: Value = Value * 1E+17
Case 18: Value = Value * 1E+18
Case 19: Value = Value * 1E+19
Case 20: Value = Value * 1E+20
Case 21: Value = Value * 1E+21
Case 22: Value = Value * 1E+22
Case 23: Value = Value * 1E+23
Case 24: Value = Value * 1E+24
Case 25: Value = Value * 1E+25
End Select
End Sub

Private Function ReturnEmptyByteArray() As Byte()
'returning an array with UBound = -1 would be nice
End Function

Private Function Unpack(Var As XSQLVAR, ByVal Coalesce As Boolean) As Variant
Dim VByte As Byte
Dim VStr As String
Dim VInt As Integer
Dim VLong As Long
Dim VCur As Currency
Dim VSng As Single
Dim VDbl As Double
'Extract Variant From Var, Return Empty On Failure
memcpy VInt, ByVal Var.sqlind, 2
If VInt < 0 Then
    If Not Coalesce Then
        Unpack = Null
    Else
        Select Case Var.sqltype And Not 1
        Case SQL_TEXT
            Unpack = Space$(Var.sqllen)
        Case SQL_VARYING
            Unpack = vbNullString
        Case SQL_SHORT
            Unpack = 0
        Case SQL_LONG, SQL_INT64
            Unpack = 0&
        Case SQL_BOOL
            Unpack = False
        Case SQL_TYPE_DATE, SQL_TYPE_TIME, SQL_TIMESTAMP
            Unpack = CDate(0)
        Case SQL_BLOB
            If Var.sqlsubtype = 1 Then
                Unpack = vbNullString
            Else
                Unpack = ReturnEmptyByteArray
            End If
        Case SQL_FLOAT, SQL_D_FLOAT
            Unpack = 0!
        Case SQL_DOUBLE
            Unpack = 0#
        Case Else
            Unpack = Null
            SetErr 8
            Debug.Print "SQL Type (" & (Var.sqltype And Not 1) & ") Não Implementado/Não Suportado/Type Desconhecido"
            Debug.Assert False
        End Select
    End If
    Exit Function
End If
Select Case Var.sqltype And Not 1
Case SQL_TEXT
    VStr = String$(Var.sqllen, 0)
    memcpy ByVal StrPtr(VStr), ByVal Var.sqldata, Var.sqllen
    LSet VStr = StrConv(VStr, vbUnicode)
    Unpack = VStr
Case SQL_VARYING
    Debug.Print Var.sqltype And Not 1
    memcpy VInt, ByVal Var.sqldata, 2
    If VInt > Var.sqllen Then
        Debug.Assert False
        VInt = Var.sqllen
    End If
    VStr = String$(VInt, 0)
    memcpy ByVal StrPtr(VStr), ByVal Var.sqldata + 2, VInt
    LSet VStr = StrConv(VStr, vbUnicode)
    Unpack = VStr
Case SQL_SHORT
    memcpy VInt, ByVal Var.sqldata, 2
    Unpack = VInt
    ApplyScale Unpack, Var.sqlscale
Case SQL_LONG
    memcpy VLong, ByVal Var.sqldata, 4
    Unpack = VLong
    ApplyScale Unpack, Var.sqlscale
Case SQL_INT64
    memcpy VLong, ByVal Var.sqldata + 4, 4
    If VLong Then
        VCur = VLong * 4294967296@
    End If
    memcpy VLong, ByVal Var.sqldata, 4 'FIXME Watch Out For Endiannes
    If VLong < 0 Then
        Unpack = VCur + VLong + 4294967296@
    ElseIf VCur = 0 Then
        Unpack = VLong
    Else
        Unpack = VCur + VLong
    End If
    ApplyScale Unpack, Var.sqlscale
Case SQL_BOOL
    memcpy VByte, ByVal Var.sqldata, 1
    If VByte Then
        Unpack = True
    Else
        Unpack = False
    End If
Case SQL_TYPE_DATE
    memcpy VLong, ByVal Var.sqldata, 4
    Unpack = DecodeDate(VLong)
Case SQL_TYPE_TIME
    memcpy VLong, ByVal Var.sqldata, 4
    Unpack = DecodeTime(VLong)
Case SQL_TIMESTAMP
    memcpy VLong, ByVal Var.sqldata, 4 'Assume Lower 4 Bytes is Date
    Unpack = DecodeDate(VLong)
    memcpy VLong, ByVal Var.sqldata + 4, 4 'And Higher 4 Bytes is Time
    Unpack = Unpack + DecodeTime(VLong)
Case SQL_BLOB
    Unpack = ReadBlob(Var.sqldata, Var.sqlsubtype = 1)
    If Not CacheBlob Then
        MaxC = -1
        Erase Cache
    End If
Case SQL_FLOAT, SQL_D_FLOAT
    memcpy VSng, ByVal Var.sqldata, 4
    Unpack = VSng
Case SQL_DOUBLE
    memcpy VDbl, ByVal Var.sqldata, 8
    Unpack = VDbl
Case Else 'SQL_ARRAY, SQL_QUAD, SQL_D_FLOAT:
    Unpack = Null
    SetErr 8
    Debug.Print "SQL Type (" & (Var.sqltype And Not 1) & ") Não Implementado/Não Suportado/Type Desconhecido"
    Debug.Assert False
End Select
End Function

Private Function DecodeDate(ByVal nday As Long) As Date
'ref: https://github.com/nakagami/firebirdsql/blob/master/xsqlvar.go#L226
'ref: https://github.com/FirebirdSQL/firebird/blob/master/src/common/classes/NoThrowTimeStamp.cpp#L178
Dim Century As Long
Dim Day As Long
Dim Month As Long
Dim Year As Long

nday = nday + 678882 '2400001 - 1721119
Century = (4 * nday - 1) \ 146097
nday = 4 * nday - 1 - 146097 * Century
Day = nday \ 4

nday = (4 * Day + 3) \ 1461
Day = 4 * Day + 3 - 1461 * nday
Day = (Day + 4) \ 4

Month = (5 * Day - 3) \ 153
Day = 5 * Day - 3 - 153 * Month
Day = (Day + 5) \ 5
Year = 100 * Century + nday
If Month < 10 Then
    Month = Month + 3
Else
    Month = Month - 9
    Year = Year + 1
End If
DecodeDate = DateSerial(Year, Month, Day)
End Function

Private Function DecodeTime(ByVal N As Long) As Date
'ref: https://github.com/nakagami/firebirdsql/blob/master/xsqlvar.go#L249
N = N \ 10000& 'Ignore nano seconds
DecodeTime = TimeSerial(N \ 3600&, (N \ 60&) Mod 60&, N Mod 60&)
End Function

Private Function EncodeDate(ByVal D As Date) As Long
'ref: https://github.com/FirebirdSQL/firebird/blob/master/src/common/classes/NoThrowTimeStamp.cpp#L233
Dim Month As Long, Year As Long
Month = DateTime.Month(D)
Year = DateTime.Year(D)
If Month > 2 Then
    Month = Month - 3
Else
    Month = Month + 3
    Year = Year - 1
End If
EncodeDate = ((146097 * (Year \ 100&)) \ 4& + _
             (1461& * (Year - 100& * (Year \ 100&))) \ 4& + _
             (153& * Month + 2) \ 5 + CLng(DateTime.Day(D)) - 678882)
End Function

Private Function EncodeTime(ByVal D As Date) As Long
EncodeTime = (CLng(Second(D)) + CLng(Minute(D)) * 60& + CLng(Hour(D)) * 3600&) * 10000
End Function

Private Function ReadBlob(ByVal QuadPtr As Long, ByVal Convert As Boolean) As Variant
Dim Quad(1) As Long
Dim blob_handle As Long
Dim Z As Long, Ret As Long
Dim bytesRead As Integer
memcpy Quad(0), ByVal QuadPtr, 8
For Z = 0 To MaxC
    If Cache(Z).Quad(0) = Quad(0) And Cache(Z).Quad(1) = Quad(1) Then
        GoTo return_blob
        Exit Function
    End If
    If Cache(Z).Quad(0) = 0 And Cache(Z).Quad(1) = 0 Then GoTo clear
Next
MaxC = MaxC + 1
ReDim Preserve Cache(MaxC)
clear:
Cache(Z).Quad(0) = Quad(0)
Cache(Z).Quad(1) = Quad(1)
Cache(Z).Max = -1
Const BuffLen As Integer = 16 '1024 * 4
Dim Buff(BuffLen - 1) As Byte
With Cache(Z)
    If isc_open_blob2(status_vector(0), db_instance.Handle, tr_handle, blob_handle, ByVal QuadPtr, 0, ByVal 0&) Then
        .Quad(0) = 0: .Quad(1) = 0: SetErr
        Exit Function
    End If
    Do
        Ret = isc_get_segment(status_vector(0), blob_handle, bytesRead, BuffLen, Buff(0))
        If Ret = 335544366 Then
            'There Is Still More Data To Come
            ReDim Preserve .B(.Max + bytesRead)
            memcpy .B(.Max + 1), Buff(0), bytesRead
            .Max = .Max + bytesRead
        ElseIf Ret = 0 Then
            'We Are Done Reading, Last Data Is In Buff
            ReDim Preserve .B(.Max + bytesRead)
            memcpy .B(.Max + 1), Buff(0), bytesRead
            .Max = .Max + bytesRead
            Exit Do
        Else
            .Quad(0) = 0: .Quad(1) = 0: SetErr
            Erase .B
            Z = -1
            Exit Do
        End If
    Loop
End With
If isc_close_blob(status_vector(0), blob_handle) Then SetErr
return_blob:
If Z = -1 Then
    ReadBlob = Null
ElseIf Not Convert Then
    ReadBlob = Cache(Z).B
Else
    ReadBlob = StrConv(Cache(Z).B, vbUnicode)
End If
End Function

Private Function WriteBlob(Arr As Variant, ByVal QuadPtr As Long) As Boolean
Dim blob_handle As Long
Dim Z As Long
Dim Ptr As Long
Dim PtrEnd As Long

If VarType(Arr) <> (vbArray Or vbByte) Then Exit Function

If UBound(Arr) < LBound(Arr) Then Exit Function

Ptr = VarPtr(Arr) 'Ptr = VARIANT*
memcpy Ptr, ByVal Ptr + 8, 4 'Ptr = VARIANT->SAFEARRAY*
memcpy Ptr, ByVal Ptr + 12, 4 'Ptr = VARIANT->SAFEARRAY->Data*

PtrEnd = Ptr + UBound(Arr) - LBound(Arr) + 1

If isc_create_blob2(status_vector(0), db_instance.Handle, tr_handle, blob_handle, ByVal QuadPtr, 0, ByVal 0&) Then SetErr: Exit Function
Do While PtrEnd - Ptr > &H7FFF&
    If isc_put_segment(status_vector(0), blob_handle, &H7FFF, ByVal Ptr) Then GoTo cancel
    Ptr = Ptr + &H7FFF&
Loop
If isc_put_segment(status_vector(0), blob_handle, PtrEnd - Ptr, ByVal Ptr) Then GoTo cancel
If isc_close_blob(status_vector(0), blob_handle) Then SetErr
WriteBlob = True
Exit Function
cancel: SetErr
If isc_cancel_blob(status_vector(0), blob_handle) Then SetErr
End Function


#If DEBUGING Then 'If Debugging
Public Property Get MaxXSQLVAR(B() As Byte) As Long
Dim X As XSQLDA
LoadXSQLDA B, X 'Não a coisa mais eficiente do mundo mas serve
MaxXSQLVAR = X.sqln
End Property

Public Sub LoadXSQLVAR(B() As Byte, ByVal Index As Long, Dest As XSQLVAR)
memcpy Dest, ByVal VarPtr(B(0)) + LenXSQLDA + LenXSQLVAR * Index, LenXSQLVAR
End Sub

Public Sub StoreXSQLVAR(B() As Byte, ByVal Index As Long, Src As XSQLVAR)
memcpy ByVal VarPtr(B(0)) + LenXSQLDA + LenXSQLVAR * Index, Src, LenXSQLVAR
End Sub

Public Sub LoadXSQLDA(B() As Byte, Dest As XSQLDA)
memcpy Dest, ByVal VarPtr(B(0)), LenXSQLDA
End Sub

Public Sub StoreXSQLDA(B() As Byte, Src As XSQLDA)
memcpy ByVal VarPtr(B(0)), Src, LenXSQLDA
End Sub

Public Sub Dbg(B() As Byte)
Dim Var As XSQLVAR
Dim Z As Long
For Z = 0 To MaxXSQLVAR(B) - 1
    Debug.Print "XSQLVAR #" & Z
    LoadXSQLVAR B, Z, Var
    DbgVar Var
Next
End Sub

Public Sub DbgVar(X As XSQLVAR)
TabPrint "sqltype: ", "(" & X.sqltype & ")" & DbgHex(VarPtr(X.sqltype), LenB(X.sqltype))
TabPrint "sqlscale: ", "(" & X.sqlscale & ")" & DbgHex(VarPtr(X.sqlscale), LenB(X.sqlscale))
'TabPrint "sqlprecision: ", "(" & X.sqlprecision & ")" & DbgHex(VarPtr(X.sqlprecision), LenB(X.sqlprecision))
TabPrint "sqlsubtype: ", "(" & X.sqlsubtype & ")" & DbgHex(VarPtr(X.sqlsubtype), LenB(X.sqlsubtype))
TabPrint "sqllen: ", "(" & X.sqllen & ")" & DbgHex(VarPtr(X.sqllen), LenB(X.sqllen))
TabPrint "sqldata: ", "(" & X.sqldata & ")" & DbgHex(VarPtr(X.sqldata), LenB(X.sqldata))
TabPrint "sqlind: ", "(" & X.sqlind & ")" & DbgHex(VarPtr(X.sqlind), LenB(X.sqlind))
TabPrint "sqlname_len: ", "(" & X.sqlname_len & ")" & DbgHex(VarPtr(X.sqlname_len), LenB(X.sqlname_len))
TabPrint "sqlname: ", DbgStr(X.sqlname) & DbgHex(VarPtr(X.sqlname(0)), METADATALENGTH)
TabPrint "relname_len: ", "(" & X.relname_len & ")" & DbgHex(VarPtr(X.relname_len), LenB(X.relname_len))
TabPrint "relname: ", DbgStr(X.relname) & DbgHex(VarPtr(X.relname(0)), METADATALENGTH)
TabPrint "ownname_len: ", "(" & X.ownname_len & ")" & DbgHex(VarPtr(X.ownname_len), LenB(X.ownname_len))
TabPrint "ownname: ", DbgStr(X.ownname) & DbgHex(VarPtr(X.ownname(0)), METADATALENGTH)
TabPrint "aliasname_len: ", "(" & X.aliasname_len & ")" & DbgHex(VarPtr(X.aliasname_len), LenB(X.aliasname_len))
TabPrint "aliasname: ", DbgStr(X.aliasname) & DbgHex(VarPtr(X.aliasname(0)), METADATALENGTH)
End Sub

Public Sub TabPrint(A As String, B As String)
If Len(A) > 16 Then
    Debug.Print A & B
Else
    Debug.Print A & Space$(16 - Len(A)) & B
End If
End Sub

Public Function DbgStr(Str() As Byte) As String
Dim Z As Long
DbgStr = StrConv(Str, vbUnicode)
For Z = 0 To UBound(Str)
    If Str(Z) = 0 Then
        DbgStr = Left$(DbgStr, Z)
        Exit For
    End If
Next
DbgStr = """" & DbgStr & """"
End Function

Public Function DbgHex(ByVal Ptr As Long, ByVal Length As Long) As String
Dim B As Byte
Dim Begin As Long
Begin = Ptr
Do While Length > 0
    memcpy B, ByVal Ptr, 1
    DbgHex = DbgHex & " "
    If B \ &H10 >= 10 Then
        DbgHex = DbgHex & Chr$(B \ &H10 + 55)
    Else
        DbgHex = DbgHex & Chr$(B \ &H10 + 48)
    End If
    B = B And &HF
    If B >= 10 Then
        DbgHex = DbgHex & Chr$(B + 55)
    Else
        DbgHex = DbgHex & Chr$(B + 48)
    End If
    Length = Length - 1
    Ptr = Ptr + 1
Loop
End Function
#End If


