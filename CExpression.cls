VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum EConst
    Begin_
    FieldMark_L '[
    FieldMark_R ']
    FieldMark_Func '_F
    ConstMark_L '"
    ConstMark_R '"
    ConstMark_Func '_C
    FuncMark_L '(
    FuncMark_R ')
    ArgumentMask '___agmt___
    ErrMissFieldMark_L 'miss [
    ErrMissFieldMark_R 'miss ]
    ErrMissConstMark_L 'miss "
    ErrMissConstMark_R 'miss "
    ErrMissFuncMark_L 'miss (
    ErrMissFuncMark_R 'miss )
    End_
End Enum

Private m_result As Long

Private m_arguments()
Private m_argumentCount As Long

Private m_errDesc As String

Private Sub Class_Initialize()
    m_argumentCount = 0
End Sub

Private Sub Class_Terminate()
    ''''
End Sub

Public Property Get Result() As Long
    Result = m_result
End Property

Public Property Get ErrDesc() As Variant
    ErrDesc = m_errDesc
End Property

Private Function Con(ByVal ConId As EConst)
    Static consts As Variant
    If Not IsArray(consts) Then
        Dim arr(EConst.Begin_ + 1 To EConst.End_ - 1)
        arr(EConst.FieldMark_L) = "["
        arr(EConst.FieldMark_R) = "]"
        arr(EConst.FieldMark_Func) = "_F"
        arr(EConst.ConstMark_L) = """"
        arr(EConst.ConstMark_R) = """"
        arr(EConst.ConstMark_Func) = "_C"
        arr(EConst.FuncMark_L) = "("
        arr(EConst.FuncMark_R) = ")"
        arr(EConst.ArgumentMask) = "{_Argument_}"
        arr(EConst.ErrMissFieldMark_L) = "miss '['"
        arr(EConst.ErrMissFieldMark_R) = "miss ']'"
        arr(EConst.ErrMissConstMark_L) = "miss '""'"
        arr(EConst.ErrMissConstMark_R) = "miss '""'"
        arr(EConst.ErrMissFuncMark_L) = "miss '('"
        arr(EConst.ErrMissFuncMark_R) = "miss ')'"
        consts = arr
    End If
    
    Con = consts(ConId)
End Function

Private Function appendArgument(Optional ByVal body_ As String = "", _
    Optional ByVal Type_ As EAType = EAType.Const_, _
    Optional ByVal funcAddr_ As Long = -1, _
    Optional ByVal funcName_ As String = "", _
    Optional ByVal arguCount_ As Long = -1, _
    Optional ByVal value_ As Variant = Empty) As Variant
    
    Dim vArgu: vArgu = NewExpArgument
    
    vArgu(ExpArgu.Id) = m_argumentCount
    vArgu(ExpArgu.Body) = body_
    vArgu(ExpArgu.MaskBody) = Con(ArgumentMask) & m_argumentCount
    vArgu(ExpArgu.Type_) = Type_
    vArgu(ExpArgu.FuncAddr) = funcAddr_
    vArgu(ExpArgu.funcName) = funcName_
    vArgu(ExpArgu.Value) = value_
    vArgu(ExpArgu.ArguCount) = arguCount_
    
    If m_argumentCount = 0 Then
        ReDim m_arguments(0)
    Else
        ReDim Preserve m_arguments(m_argumentCount)
    End If
    m_arguments(m_argumentCount) = vArgu
    m_argumentCount = m_argumentCount + 1
    
    appendArgument = vArgu
End Function

Private Function appendArgumentEx(ByRef vArgument As Variant)
    vArgument(ExpArgu.Id) = m_argumentCount
    vArgument(ExpArgu.MaskBody) = Con(ArgumentMask) & m_argumentCount
    
    If m_argumentCount = 0 Then
        ReDim m_arguments(0)
    Else
        ReDim Preserve m_arguments(m_argumentCount)
    End If
    m_arguments(m_argumentCount) = vArgument
    m_argumentCount = m_argumentCount + 1
    
    appendArgumentEx = vArgument
End Function

'Private Function updateArgument(ByVal Id As Long, _
'    Optional ByVal body_ As Variant = Empty, _
'    Optional ByVal Type_ As Variant = Empty, _
'    Optional ByVal funcAddr_ As Variant = Empty, _
'    Optional ByVal funcName_ As Variant = Empty, _
'    Optional ByVal arguments_ As Variant = Empty, _
'    Optional ByVal arguCount_ As Variant = Empty, _
'    Optional ByVal value_ As Variant = Empty) As Variant
'    If Id < m_argumentCount Or Id >= m_argumentCount Then Exit Function
'    If Not IsEmpty(body_) Then m_arguments(Id)(ExpArgu.Body) = body_
'    If Not IsEmpty(Type_) Then m_arguments(Id)(ExpArgu.Type_) = Type_
'    If Not IsEmpty(funcAddr_) Then m_arguments(Id)(ExpArgu.FuncAddr) = funcAddr_
'    If Not IsEmpty(funcName_) Then m_arguments(Id)(ExpArgu.funcName) = funcName_
'    If Not IsEmpty(arguCount_) Then m_arguments(Id)(ExpArgu.ArguCount) = arguCount_
'    If Not IsEmpty(Arguments) Then m_arguments(Id)(ExpArgu.Arguments) = Arguments
'    If Not IsEmpty(value_) Then m_arguments(Id)(ExpArgu.Value) = value_
'    updateArgument = m_arguments(Id)
'End Function
'
'Private Function updateArgumentEx(ByVal Id As Long, ByVal Property As ExpArgu, ByVal NewValue As Variant) As Variant
'    If Id < m_argumentCount Or Id >= m_argumentCount Then Exit Function
'    m_arguments(Id)(Property) = NewValue
'    updateArgumentEx = m_arguments(Id)
'End Function

Private Function parseArgumentExp(ByVal Expression As String)
    Dim chrPos As Long: chrPos = InStr(1, Expression, Con(ArgumentMask))
    
    If chrPos > 0 Then
        Dim ArguId As Long: ArguId = CLng(Mid$(Expression, chrPos + Len(Con(ArgumentMask))))
        Dim arguFunc As Variant: arguFunc = m_arguments(ArguId)
        Dim ret As Variant
        
        If arguFunc(ExpArgu.Type_) = EAType.func Then
            Dim funcName As String
            Dim Arguments As Variant
            Dim arguExp As String
            Dim i As Long
            
            funcName = Trim$(Left$(Expression, chrPos - 1))
            arguExp = Trim$(arguFunc(ExpArgu.Body))
            If funcName = "" Then
                If (Left$(arguExp, 1) = Con(FuncMark_L) And Right$(arguExp, 1) = Con(FuncMark_R)) Then
                    ret = doParse(Mid$(arguExp, 2, Len(arguExp) - 2))
                Else
                    ret = doParse(arguExp)
                End If
            Else
                Dim f As Variant: f = MExpFuncs.GetExpFuncByName(funcName)
                If IsArray(f) Then
                    arguFunc(ExpArgu.funcName) = f(ExpArgu.funcName)
                    arguFunc(ExpArgu.FuncAddr) = f(ExpArgu.FuncAddr)
                    arguFunc(ExpArgu.ArguCount) = f(ExpArgu.ArguCount)
                    
                    arguExp = Trim$(Mid$(arguExp, 2, Len(arguExp) - 2))
                    
                    If Len(arguExp) > 0 Then
                        Arguments = Split(arguExp, ",")
                        For i = 0 To UBound(Arguments)
                            Arguments(i) = doParse(Trim$(Arguments(i)))(ExpArgu.Id)
                        Next
                        arguFunc(ExpArgu.Arguments) = Arguments
                    Else
                        If arguFunc(ExpArgu.ArguCount) > 0 Then
                            m_errDesc = "miss argument(s) for function '" & funcName & "'"
                        End If
                    End If
                    
                    m_arguments(ArguId) = arguFunc
                    ret = arguFunc
                Else
                    arguFunc(ExpArgu.Type_) = EAType.Const_
                    arguFunc(ExpArgu.Value) = arguFunc(ExpArgu.Body)
                    
                    m_arguments(ArguId) = arguFunc
                    ret = arguFunc
                    
                    m_errDesc = "unknow function '" & funcName & "'"
                End If
            End If
        Else
            ret = arguFunc
        End If
    Else
        ret = appendArgument(body_:=Expression, value_:=Trim$(Expression))
    End If
    
    parseArgumentExp = ret
End Function

Private Function parseMinusExp(ByVal Expression As String) ', ByVal Operator As String, ByVal OperatorFunc As ExpFunc
    Dim argu As Variant
    argu = MExpFuncs.GetExpFunc(ExpFunc.f_minus)
    argu(ExpArgu.Body) = Expression
    
    Dim arr As Variant: arr = Split(Expression, EFN_MINUS)
    Dim n As Long: n = UBound(arr)
    Dim Arguments() As Long: ReDim Arguments(n) As Long
    Dim i As Long
    
    For i = 0 To n
        Arguments(i) = doParse(Trim$(arr(i)))(ExpArgu.Id)
    Next
    
    argu(ExpArgu.Arguments) = Arguments

    parseMinusExp = appendArgumentEx(argu)
End Function

Private Function parseOperatorExp(ByVal Expression As String, ByVal Operator As String, ByVal OperatorFunc As ExpFunc)
    Dim argu As Variant
    argu = MExpFuncs.GetExpFunc(OperatorFunc)
    argu(ExpArgu.Body) = Expression
    
    Dim arr As Variant: arr = Split(Expression, Operator)
    Dim n As Long: n = UBound(arr)
    Dim Arguments() As Long: ReDim Arguments(n) As Long
    Dim i As Long
    
    For i = 0 To n
        Arguments(i) = doParse(Trim$(arr(i)))(ExpArgu.Id)
    Next
    
    argu(ExpArgu.Arguments) = Arguments

    parseOperatorExp = appendArgumentEx(argu)
End Function

Private Function doParse(ByVal Expression As String)
    Dim ret
    
    If InStr(1, Expression, EFN_JOIN_STR) > 0 Then
        ret = parseOperatorExp(Expression, EFN_JOIN_STR, ExpFunc.f_join_str)
    ElseIf InStr(1, Expression, EFN_PLUS) > 0 Then
        ret = parseOperatorExp(Expression, EFN_PLUS, ExpFunc.f_plus)
    ElseIf InStr(1, Expression, EFN_MINUS) > 0 Then
        ret = parseMinusExp(Expression)
    ElseIf InStr(1, Expression, EFN_MULTIPLY) > 0 Then
        ret = parseOperatorExp(Expression, EFN_MULTIPLY, ExpFunc.f_multiply)
    ElseIf InStr(1, Expression, EFN_DIVIDE) > 0 Then
        ret = parseOperatorExp(Expression, EFN_DIVIDE, ExpFunc.f_divide)
    ElseIf InStr(1, Expression, EFN_SMALLER) > 0 Then
        ret = parseOperatorExp(Expression, EFN_SMALLER, ExpFunc.f_smaller)
    ElseIf InStr(1, Expression, EFN_LARGER) > 0 Then
        ret = parseOperatorExp(Expression, EFN_LARGER, ExpFunc.f_larger)
    ElseIf InStr(1, Expression, EFN_EQUAL) > 0 Then
        ret = parseOperatorExp(Expression, EFN_EQUAL, ExpFunc.f_equal)
    Else
        ret = parseArgumentExp(Expression)
    End If
    
    doParse = ret
End Function

Private Function maskConsts(ByVal Expression As String) As String
    Dim exp_len As Long: exp_len = Len(Expression)
    Dim u_exp As String: u_exp = UCase(Expression)
    Dim vArg As Variant
    Dim iChr As Long: iChr = 1
    Dim chrPos As Long
    Dim chr2Pos As Long
    Dim nConstFunc As Long: nConstFunc = Len(Con(ConstMark_Func))
    Dim nFieldFunc As Long: nFieldFunc = Len(Con(FieldMark_Func))
    Dim strChr As String
    Dim ret() As String: ReDim ret(1 To exp_len) As String
    Dim partExp As String
    Dim constVal As String
    
    Do
        If Len(m_errDesc) > 0 Then Exit Do
        If iChr > exp_len Then Exit Do
        
        If Mid(u_exp, iChr, 1) = Con(ConstMark_L) Then
            '' const value mark by ""
            chrPos = InStr(iChr + 1, u_exp, Con(ConstMark_R))
            If chrPos > iChr Then
                partExp = Mid$(Expression, iChr, chrPos - iChr + 1)
                constVal = Trim$(Mid$(Expression, iChr + 1, chrPos - iChr - 1))
                vArg = appendArgument(body_:=partExp, value_:=constVal)
                ret(iChr) = vArg(ExpArgu.MaskBody)
                iChr = chrPos + 1
            Else
                m_errDesc = Con(ErrMissConstMark_R)
            End If
        ElseIf StrComp(Mid$(u_exp, iChr, nConstFunc), Con(ConstMark_Func), vbTextCompare) = 0 Then
            '' const value mark by _C("")
            chrPos = iChr + nConstFunc
            If MFunc.NextTrimChar(u_exp, chrPos, exp_len) = Con(FuncMark_L) Then
                chr2Pos = InStr(chrPos, u_exp, Con(FuncMark_R))
                If chr2Pos > chrPos Then
                    partExp = Mid$(Expression, iChr, chr2Pos - iChr + 1)
                    constVal = Trim$(Mid$(Expression, chrPos + 1, chr2Pos - chrPos - 1))
                    vArg = appendArgument(body_:=partExp, value_:=constVal)
                    ret(iChr) = vArg(ExpArgu.MaskBody)
                    iChr = chr2Pos + 1
                Else
                    m_errDesc = Con(ErrMissFuncMark_R)
                End If
            Else
                ret(iChr) = Mid(Expression, iChr, 1)
                iChr = iChr + 1
            End If
        ElseIf Mid(u_exp, iChr, 1) = Con(FieldMark_L) Then
            '' const field mark by []
            chrPos = InStr(iChr + 1, u_exp, Con(FieldMark_R))
            If chrPos > iChr Then
                partExp = Mid$(Expression, iChr, chrPos - iChr + 1)
                constVal = Trim$(Mid$(Expression, iChr + 1, chrPos - iChr - 1))
                vArg = appendArgument(body_:=partExp, Type_:=EAType.Field, value_:=constVal)
                ret(iChr) = vArg(ExpArgu.MaskBody)
                iChr = chrPos + 1
            Else
                m_errDesc = Con(ErrMissFuncMark_R)
            End If
        ElseIf StrComp(Mid$(u_exp, iChr, nFieldFunc), Con(FieldMark_Func), vbTextCompare) = 0 Then
            '' const field mark by _F("")
            chrPos = iChr + nFieldFunc
            If MFunc.NextTrimChar(u_exp, chrPos, exp_len) = Con(FuncMark_L) Then
                chr2Pos = InStr(chrPos, u_exp, Con(FuncMark_R))
                If chr2Pos > chrPos Then
                    partExp = Mid$(Expression, iChr, chr2Pos - iChr + 1)
                    constVal = Trim$(Mid$(Expression, chrPos + 1, chr2Pos - chrPos - 1))
                    vArg = appendArgument(body_:=partExp, Type_:=EAType.Field, value_:=constVal)
                    ret(iChr) = vArg(ExpArgu.MaskBody)
                    iChr = chr2Pos + 1
                Else
                    m_errDesc = Con(ErrMissFuncMark_R)
                End If
            Else
                ret(iChr) = Mid(Expression, iChr, 1)
                iChr = iChr + 1
            End If
        Else
            ret(iChr) = Mid(Expression, iChr, 1)
            iChr = iChr + 1
        End If
    Loop
    
    maskConsts = Join(ret, "")
End Function

Private Function maskFunctions(ByVal Expression As String) As String
    Dim exp_len As Long
    Dim vArg As Variant
    Dim iChr As Long: iChr = 1
    Dim strChr As String
    Dim iFunc As Long: iFunc = 0
    Dim nFunc As Long: nFunc = 0
    Dim posPair() As Long
    Dim partExp As String
    
    Do
        exp_len = Len(Expression)
        If iChr > exp_len Then Exit Do
        
        strChr = Mid$(Expression, iChr, 1)
        
        If strChr = Con(FuncMark_L) Then
            iFunc = iFunc + 1
            ReDim Preserve posPair(1 To iFunc) As Long
            posPair(iFunc) = iChr
            iChr = iChr + 1
        ElseIf strChr = Con(FuncMark_R) Then
            If iFunc < 1 Then
                m_errDesc = Con(ErrMissFuncMark_L)
                Exit Do
            Else
                partExp = Mid$(Expression, posPair(iFunc), iChr - posPair(iFunc) + 1)
                vArg = appendArgument(body_:=partExp, Type_:=EAType.func)
                Expression = Left$(Expression, posPair(iFunc) - 1) & vArg(ExpArgu.MaskBody) & Mid$(Expression, iChr + 1)
                iChr = posPair(iFunc) + Len(vArg(ExpArgu.MaskBody)) - 1
                iFunc = iFunc - 1
            End If
        Else
            iChr = iChr + 1
        End If
    Loop
    
    If iFunc > 0 Then m_errDesc = Con(ErrMissFuncMark_L)
    
    maskFunctions = Expression
End Function

Public Function Parse(ByVal Expression As String)
    m_argumentCount = 0
    m_errDesc = ""
    
    Dim stdExp As String
    stdExp = maskConsts(Trim$(Expression))
    stdExp = maskFunctions(stdExp)
    
    m_result = doParse(stdExp)(ExpArgu.Id)
End Function

Private Function createArray(Optional ByVal DataType As Long = 1, Optional ByVal Start As Long = 0) As CArray
    Dim ret As New CArray
    ret.Type_ = DataType
    ret.StartPos = Start
    Set createArray = ret
End Function

Private Function argumentToXML(ByVal ArguId As Long) As String
    Dim oArr As CArray
    Dim Argument As Variant: Argument = m_arguments(ArguId)
    
    If IsArray(Argument) Then
        Set oArr = createArray(4)
        
        With oArr
            .AppendStrItem "<func"
            .AppendStrItem " id=""" & Argument(ExpArgu.Id) & """"
            .AppendStrItem " name=""" & CXml(Argument(ExpArgu.funcName)) & """"
            .AppendStrItem " addr=""" & Argument(ExpArgu.FuncAddr) & """"
            .AppendStrItem " value=""" & CXml(Argument(ExpArgu.Value)) & """"
            Select Case Argument(ExpArgu.Type_)
                Case EAType.Const_
                    .AppendStrItem " type=""Const"""
                Case EAType.Field
                    .AppendStrItem " type=""Field"""
                Case EAType.func
                    .AppendStrItem " type=""Function"""
                Case EAType.Unknow
                    .AppendStrItem " type=""Unknow"""
                Case EAType.Var
                    .AppendStrItem " type=""Var"""
            End Select
            .AppendStrItem " body=""" & CXml(Argument(ExpArgu.Body)) & """"
            .AppendStrItem ">"
            
            If IsArray(Argument(ExpArgu.Arguments)) Then
                Dim i As Long
                For i = 0 To UBound(Argument(ExpArgu.Arguments))
                    .AppendStrItem argumentToXML(Argument(ExpArgu.Arguments)(i))
                Next
            End If
            
            .AppendStrItem "</func>"
            
            argumentToXML = Join(.StrList, "")
        End With
    Else
        argumentToXML = Argument
    End If
End Function

Public Function ToXML() As String
    ToXML = argumentToXML(m_result)
End Function
