VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum ArrType
    Variant_ = 1
    Long_ = 2
    Double_ = 3
    String_ = 4
End Enum

Private m_type As Long

Private m_varArr() As Variant
Private m_lngArr() As Long
Private m_dblArr() As Double
Private m_strArr() As String

Private m_count As Long
Private m_start As Long

Private Sub Class_Initialize()
    m_count = 0
    m_start = 0
End Sub

Public Property Get Type_() As Long
    Type_ = m_type
End Property

Public Property Let Type_(ByVal vNewValue As Long)
    m_type = vNewValue
End Property

Public Property Get StartPos() As Long
    StartPos = m_start
End Property

Public Property Let StartPos(ByVal vNewValue As Long)
    m_start = vNewValue
End Property

Public Property Get VarList() As Variant()
    VarList = m_varArr
End Property

Public Property Get LngList() As Long()
    LngList = m_lngArr
End Property

Public Property Get DblList() As Double()
    DblList = m_dblArr
End Property

Public Property Get StrList() As String()
    StrList = m_strArr
End Property

Public Function AppendVarItem(item) As Long
    Dim appendId As Long
    If m_count = 0 Then
        appendId = m_start
        ReDim m_varArr(m_start To m_start) As Variant
    Else
        appendId = m_start + m_count
        ReDim Preserve m_varArr(m_start To appendId) As Variant
    End If
    m_varArr(appendId) = item
    m_count = m_count + 1
    AppendVarItem = appendId
End Function

Public Function AppendLngItem(item As Long) As Long
    Dim appendId As Long
    If m_count = 0 Then
        appendId = m_start
        ReDim m_lngArr(m_start To m_start) As Long
    Else
        appendId = m_start + m_count
        ReDim Preserve m_lngArr(m_start To appendId) As Long
    End If
    m_lngArr(appendId) = item
    m_count = m_count + 1
    AppendLngItem = appendId
End Function

Public Function AppendDblItem(item As Double) As Long
    Dim appendId As Long
    If m_count = 0 Then
        appendId = m_start
        ReDim m_dblArr(m_start To m_start) As Double
    Else
        appendId = m_start + m_count
        ReDim Preserve m_dblArr(m_start To appendId) As Double
    End If
    m_dblArr(appendId) = item
    m_count = m_count + 1
    AppendDblItem = appendId
End Function

Public Function AppendStrItem(item As String) As Long
    Dim appendId As Long
    If m_count = 0 Then
        appendId = m_start
        ReDim m_strArr(m_start To m_start) As String
    Else
        appendId = m_start + m_count
        ReDim Preserve m_strArr(m_start To appendId) As String
    End If
    m_strArr(appendId) = item
    m_count = m_count + 1
    AppendStrItem = appendId
End Function

Public Function AppendItem(vItem) As Long
    Select Case m_type
        Case ArrType.Variant_: AppendItem = Me.AppendVarItem(vItem)
        Case ArrType.Long_: AppendItem = Me.AppendVarItem(CLng(vItem))
        Case ArrType.Double_: AppendItem = Me.AppendVarItem(CDbl(vItem))
        Case ArrType.String_: AppendItem = Me.AppendVarItem(CStr(vItem))
    End Select
End Function

Public Function Clear()
    m_count = 0
End Function

Public Property Get Count() As Long
    Count = m_count
End Property

Public Property Get List() As Variant
    Select Case m_type
        Case ArrType.Variant_: List = m_varArr
        Case ArrType.Long_: List = m_lngArr
        Case ArrType.Double_: List = m_dblArr
        Case ArrType.String_: List = m_strArr
    End Select
End Property

Public Sub UpdateVarItem(ByVal Which As Long, vItem As Variant)
    m_varArr(Which) = vItem
End Sub

Public Sub UpdateLngItem(ByVal Which As Long, vItem As Long)
    m_lngArr(Which) = vItem
End Sub

Public Sub UpdateDblItem(ByVal Which As Long, vItem As Double)
    m_dblArr(Which) = vItem
End Sub

Public Sub UpdateStrItem(ByVal Which As Long, vItem As String)
    m_strArr(Which) = vItem
End Sub

Public Sub UpdateItem(ByVal Which As Long, vItem)
    Select Case m_type
        Case ArrType.Variant_: Me.UpdateVarItem Which, vItem
        Case ArrType.Long_: Me.UpdateLngItem Which, CLng(vItem)
        Case ArrType.Double_: Me.UpdateDblItem Which, CDbl(vItem)
        Case ArrType.String_: Me.UpdateStrItem Which, CStr(vItem)
    End Select
End Sub
