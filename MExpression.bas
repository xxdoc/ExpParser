Attribute VB_Name = "MExpression"
Option Explicit

Public Enum EAType
    Const_ 'const
    Field '
    func 'function
    Var 'variable
    Unknow '
End Enum

Public Enum ExpArgu
    Begin_
    Id
    Body
    MaskBody
    Type_
    FuncAddr
    funcName
    Arguments
    ArguCount
    Value
    End_
End Enum

Public Enum ConfigCell
    Begin_
    expType
    parseRet
    execId
    End_
End Enum

Public Function NewExpArgument()
    Dim arr(ExpArgu.Begin_ + 1 To ExpArgu.End_ - 1)
    NewExpArgument = arr
End Function

Public Function NewExpFunction(ByVal Addr As Long, ByVal Name_ As String, Optional ByVal ArgumentCount As Long = -1)
    Dim arr(ExpArgu.Begin_ + 1 To ExpArgu.End_ - 1)
    arr(ExpArgu.Type_) = EAType.func
    arr(ExpArgu.FuncAddr) = Addr
    arr(ExpArgu.funcName) = Name_
    arr(ExpArgu.ArguCount) = ArgumentCount
    NewExpFunction = arr
End Function

