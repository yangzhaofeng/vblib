Attribute VB_Name = "CopyMemory"
'CopyMemory 函数声明
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
'获得一个变量 TargetVar 的地址，相当于 C 语言中的 Get- Pointer = &TargetVar
Public Function GetPointer(ByRefTargetVarAsLong) As Long
    GetPointer = VarPtr(TargetVar)
End Function
'获得一个指针变量p的值，相当于C语言中的ValueFrom- Addr = *p
Public Function ValueFromAddr(ByVal p As Long) As Long
    Dim lngTemp As Long
    CopyMemory lngTemp, ByVal p, 4
    ValueFromAddr = lngTemp
End Function
'给指针变量 p 赋值，相当于 C 语言中的*p = NewValue
Public Sub SetValueFromAddr(ByVal p As Long, NewValue As Long)
    CopyMemory ByVal p, NewValue, 4
End Sub
'将函数地址 Address 赋值给指针变量 tVarAddr，相当于 C 语言中的 tVarAddr = Address
Public Sub LetFunctionAddress(ByValtVarAddrAsLong, ByVal Address As Long)
    CopyMemory ByVal tVarAddr, Address, 4
End Sub
