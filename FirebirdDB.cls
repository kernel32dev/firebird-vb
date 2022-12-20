VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FirebirdDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

#Const DEBUGING = False
#If DEBUGING Then
Const METADATALENGTH As Long = 30

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

Dim LenXSQLDA As Long
Dim LenXSQLVAR As Long

#End If

Private Declare Sub memcpy Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub memcls Lib "kernel32" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Private Declare Function isc_attach_database Lib "fbclient.dll" (status_vector As Long, ByVal db_name_length As Integer, ByVal db_name As String, db_handle As Long, ByVal parm_buffer_length As Integer, ByVal parm_buffer As String) As Long
Private Declare Function isc_create_database Lib "fbclient.dll" (status_vector As Long, ByVal db_name_length As Integer, ByVal db_name As String, db_handle As Long, ByVal parm_buffer_length As Integer, ByVal parm_buffer As String, Optional ByVal db_type As Integer) As Long 'db_type is unused
Private Declare Function isc_detach_database Lib "fbclient.dll" (status_vector As Long, db_handle As Long) As Long
Private Declare Function isc_drop_database Lib "fbclient.dll" (status_vector As Long, db_handle As Long) As Long
Private Declare Function isc_start_multiple Lib "fbclient.dll" (status_vector As Long, tr_handle As Long, ByVal amount_of_tebs As Integer, teb As Any) As Long
Private Declare Function isc_commit_transaction Lib "fbclient.dll" (status_vector As Long, tr_handle As Long) As Long
Private Declare Function isc_rollback_transaction Lib "fbclient.dll" (status_vector As Long, tr_handle As Long) As Long

Private Enum TpbFlagsEnum
'isc_tpb_version3 = 3
isc_tpb_consistency = 1
isc_tpb_concurrency = 2
isc_tpb_shared = 3
isc_tpb_protected = 4
isc_tpb_exclusive = 5
isc_tpb_wait = 6
isc_tpb_nowait = 7
isc_tpb_read = 8
isc_tpb_write = 9
isc_tpb_lock_read = 10 'HAS ARG Chr$(3),"P","R","O"
isc_tpb_lock_write = 11 'HAS ARG Chr$(3),"P","R","O"
isc_tpb_verb_time = 12
isc_tpb_commit_time = 13
isc_tpb_ignore_limbo = 14
isc_tpb_read_committed = 15
isc_tpb_autocommit = 16
isc_tpb_rec_version = 17
isc_tpb_no_rec_version = 18
isc_tpb_restart_requests = 19
isc_tpb_no_auto_undo = 20
isc_tpb_lock_timeout = 21
isc_tpb_read_consistency 22
isc_tpb_at_snapshot_number 23
End Enum

Private Enum TransactionIsolationLevel
TIL_Concurrency = 1 'isc_tpb_concurrency
TIL_Exclusive = 2 'isc_tpb_exclusive
TIL_Consistency = 3 'isc_tpb_consistency
TIL_Read_Committed = 4 'isc_tpb_read_committed isc_tpb_rec_version
TIL_Read_Committed_Latest = 5 'isc_tpb_read_committed isc_tpb_no_rec_version
End Enum

Dim db_handle As Long
Dim db_database As String

Private Type Slot
    Handle As Long
    C() As Long 'Firebird ObjPtr
    Max As Long
End Type

Dim C() As Slot
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

Property Get Handle() As Long
Handle = db_handle
End Property

Property Get Location() As Long
Location = db_database
End Property

Function Create(Database As String, Optional Username As String = "SYSDBA", Optional Password As String = "masterke") As Boolean
Create = Connect(True, Database, Username, Password)
End Function

Function Attach(Database As String, Optional Username As String = "SYSDBA", Optional Password As String = "masterke") As Boolean
Attach = Connect(False, Database, Username, Password)
End Function

Private Function Connect(ByVal Create As Boolean, Database As String, Optional Username As String = "SYSDBA", Optional Password As String = "masterke") As Boolean
Const isc_dpb_version As Long = 1
Const isc_dpb_sql_dialect As Long = 63
Const isc_dpb_user_name As Long = 28
Const isc_dpb_password As Long = 29
#If DEBUGING Then
Dim TmpDA As XSQLDA
Dim TmpVAr As XSQLVAR
LenXSQLDA = LenB(TmpDA)
LenXSQLVAR = LenB(TmpVAr)
#End If
If db_handle Then
    Detach
Else
    MaxC = -1
End If
Dim Parm As String
Parm = Chr$(isc_dpb_version) & _
Chr$(isc_dpb_user_name) & Chr$(Len(Username)) & Username & _
Chr$(isc_dpb_password) & Chr$(Len(Password)) & Password & _
Chr$(isc_dpb_sql_dialect) & Chr$(3) & Chr$(0) & Chr$(0) & Chr$(0)
If Create Then
Connect = 0 = isc_create_database(status_vector(0), Len(Database), Database, db_handle, Len(Parm), Parm)
Else
Connect = 0 = isc_attach_database(status_vector(0), Len(Database), Database, db_handle, Len(Parm), Parm)
End If
If Not Connect Then
    SetErr
Else
    db_database = Database
    z_FBAD Me
End If
End Function

'Will detach all instances even on a fail
Function Drop() As Boolean
If db_handle = 0 Then SetErr 3: Exit Function
DetachInstances 0
If isc_drop_database(status_vector(0), db_handle) Then
    SetErr
Else
    db_database = vbNullString
    db_handle = 0
    z_FBAD Me
    Drop = True
End If
End Function

Sub Detach()
If db_handle = 0 Then Exit Sub
db_database = vbNullString
DetachInstances 0
If isc_detach_database(status_vector(0), db_handle) Then SetErr
db_handle = 0
End Sub

Function Start(Optional ByVal ReadOnly As Boolean = False) As Long
If db_handle = 0 Then SetErr 3: Exit Function
Dim Z(2) As Long
Dim Parm(1) As Byte
Parm(0) = 3 'isc_tpb_version3
If ReadOnly Then Parm(1) = isc_tpb_read Else Parm(1) = isc_tpb_write
Z(0) = VarPtr(db_handle)
Z(1) = 2
Z(2) = VarPtr(Parm(0))
If isc_start_multiple(status_vector(0), Start, 1, Z(0)) Then
    Start = 0
    SetErr
End If
End Function

Function Commit(ByVal tr_handle As Long) As Boolean
If db_handle = 0 Then SetErr 3: Exit Function
If tr_handle = 0 Then SetErr 1: Exit Function
DetachInstances tr_handle
Commit = 0 = isc_commit_transaction(status_vector(0), tr_handle)
If Not Commit Then SetErr
End Function

Function Rollback(ByVal tr_handle As Long) As Boolean
If db_handle = 0 Then SetErr 3: Exit Function
If tr_handle = 0 Then SetErr 1: Exit Function
DetachInstances tr_handle
Rollback = 0 = isc_rollback_transaction(status_vector(0), tr_handle)
If Not Rollback Then SetErr
End Function

Function Forget(ByVal tr_handle As Long) As Boolean
If CommitByDefault Then
    Forget = Commit(tr_handle)
Else
    Forget = Rollback(tr_handle)
End If
End Function

Function Execute(SqlStatement As String, Optional ByVal Transaction As Long) As Boolean
With New Firebird
    Execute = .Execute(SqlStatement, Transaction, Me)
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteCount(SqlStatement As String, Optional ByVal Transaction As Long) As Long
ExecuteCount = -1
With New Firebird
    If .Execute("SELECT COUNT(*) FROM ( " & SqlStatement & " )", Transaction, Me, False) Then
        If .Fetch Then ExecuteCount = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteFetch(SqlStatement As String, Optional ByVal Transaction As Long) As Variant
With New Firebird
    If .Execute(SqlStatement, Transaction, Me, False) Then
        If .Fetch Then ExecuteFetch = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteGenId(GeneratorName As String, Optional ByVal Increment As Long = 1, Optional ByVal Transaction As Long) As Long
ExecuteGenId = -1
With New Firebird
    If .Execute("SELECT GEN_ID(" & GeneratorName & "," & Increment & ") FROM RDB$DATABASE", Transaction, Me, False) Then
        If .Fetch Then ExecuteGenId = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteGetTime(Optional ByVal Transaction As Long) As Date
If Transaction = 0 Then Transaction = Start(True)
With New Firebird
    If .Execute("SELECT CURRENT_TIME FROM RDB$DATABASE", Transaction, Me, False) Then
        If .Fetch Then ExecuteGetTime = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteGetDate(Optional ByVal Transaction As Long) As Date
If Transaction = 0 Then Transaction = Start(True)
With New Firebird
    If .Execute("SELECT CURRENT_DATE FROM RDB$DATABASE", Transaction, Me, False) Then
        If .Fetch Then ExecuteGetDate = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Function ExecuteGetDateTime(Optional ByVal Transaction As Long) As Date
If Transaction = 0 Then Transaction = Start(True)
With New Firebird
    If .Execute("SELECT CURRENT_TIMESTAMP FROM RDB$DATABASE", Transaction, Me, False) Then
        If .Fetch Then ExecuteGetDateTime = .RowI(0)
    End If
    If .Err Then LastErr = .LastErr
    .Finish
End With
End Function

Sub z_FBTA(ByVal tr_handle As Long, ByVal Inst As Long, ByVal Listen As Boolean)
Dim Z As Long
Dim X As Long
For Z = 0 To MaxC
    If C(Z).Handle = tr_handle Then GoTo found
Next
MaxC = MaxC + 1
ReDim Preserve C(MaxC)
C(Z).Handle = tr_handle
ReDim C(Z).C(0)
C(Z).C(0) = Inst
Exit Sub
found:
With C(Z)
    If Listen Then
        .Max = .Max + 1
        ReDim Preserve .C(.Max)
        .C(.Max) = Inst
    Else
        For X = 0 To .Max
            If .C(X) = Inst Then
                If .Max = 0 Then
                    GoTo delete
                Else
                    .C(X) = .C(.Max)
                    .Max = .Max - 1
                    ReDim Preserve .C(.Max)
                    Exit Sub
                End If
            End If
        Next
    End If
    Exit Sub
delete:
End With
If MaxC = 0 Then
    MaxC = -1
    Erase C
Else
    C(Z) = C(MaxC)
    MaxC = MaxC - 1
    ReDim Preserve C(MaxC)
End If
If CommitByDefault Then
    If isc_commit_transaction(status_vector(0), tr_handle) Then SetErr
Else
    If isc_rollback_transaction(status_vector(0), tr_handle) Then SetErr
End If
End Sub

Private Sub DetachInstances(ByVal tr_handle As Long)
Dim Z As Long
Dim X As Long
Dim Firebird As Firebird
If tr_handle = 0 Then
    For Z = 0 To MaxC
        For X = 0 To C(Z).Max
            'C(Z).C(X)->Prepare vbNullString, ObjPtr(Me)
            memcpy Firebird, C(Z).C(X), 4
            Firebird.Setup -1, Me
            memcls Firebird, 4
        Next
        If CommitByDefault Then
            If isc_commit_transaction(status_vector(0), C(Z).Handle) Then SetErr
        Else
            If isc_rollback_transaction(status_vector(0), C(Z).Handle) Then SetErr
        End If
    Next
    Erase C
    MaxC = -1
Else
    For Z = 0 To MaxC
        If C(Z).Handle = tr_handle Then
            For X = 0 To C(Z).Max
                'C(Z).C(X)->Prepare vbNullString, ObjPtr(Me)
                memcpy Firebird, C(Z).C(X), 4
                Firebird.Setup -1, Nothing
                memcls Firebird, 4
            Next
            If MaxC = 0 Then
                MaxC = -1
                Erase C
            Else
                C(Z) = C(MaxC)
                MaxC = MaxC - 1
                ReDim Preserve C(MaxC)
            End If
            Exit Sub
        End If
    Next
End If
End Sub

Private Sub Class_Terminate()
Detach
End Sub
