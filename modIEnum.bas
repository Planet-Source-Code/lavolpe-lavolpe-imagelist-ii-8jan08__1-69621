Attribute VB_Name = "modIEnum"
' Borrowed from the following PSC posting
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=56800&lngWId=1
' Only modifications were to the the following:
'   - UserEnumWrapperType.UserEnum to refer to existing project class
'   - CreateEnumerator to call existing class function
'   - IEnumVariant_Next to call existing class function
'   - IEnumVariant_Skip to call existing class function
'   - IEnumVariant_Reset to call existing class function

' Original code with noted exceptions, follow

'      Date:    10/17/2004
'      Author:  Kelly Ethridge
'
'   This module creates lightweight objects that will wrap
'   a user's object that implements the IEnumerator interface.
'   By using lightweight objects, the IEnumVariant interface
'   can easily be implements, even though it is not VB friendly.
'
'   The lightweight object simply forwards the IEnumVariant calls
'   to the IEnumerable interface implemented in the user enumerator.
'
'   To learn more about lightweight objects, you should refer to
'   classic book:
'       Advanced Visual Basic 6 Power Techniques for Everyday Programs
'       By Matthew Curland

Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub VariantCopy Lib "oleaut32.dll" (ByRef pvargDest As Variant, ByRef pvargSrc As Variant)
Private Declare Function CoTaskMemAlloc Lib "ole32.dll" (ByVal cb As Long) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByRef pv As Any)
Private Declare Function IsEqualGUID Lib "ole32.dll" (ByRef rguid1 As VBGUID, ByRef rguid2 As VBGUID) As Long

Private Const E_NOINTERFACE As Long = &H80004002
Private Const BOOL_TRUE As Long = 1
Private Const BOOL_FALSE As Long = 0
Private Const ENUM_FINISHED As Long = 1

' This is the type that will wrap the user enumerator.
' When a new IEnumVariant compatible object is created,
' it will have the internal structure of UserEnumWrapperType
Private Type UserEnumWrapperType
   pVTable As Long
   cRefs As Long
   UserEnum As IItemData    ' modified by LaVolpe to use existing project Interface
End Type

' This is an array of pointers to functions that the
' object's VTable will point to.
Private Type VTable
   Functions(0 To 6) As Long
End Type

' The created VTable of function pointers
Private mVTable As VTable

' Pointer to the mVTable memory address.
Private mpVTable As Long


' GUID structure used to identify interfaces.
Private Type VBGUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

' GUIDs to identify IUnknown and IEnumVariant when
' the interface is queried.
Private IID_IUnknown As VBGUID
Private Const IID_IUnknown_Data1 As Long = 0
Private IID_IEnumVariant As VBGUID
Private Const IID_IEnumVariant_Data1 As Long = &H20404


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  Public Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Creates the LightWeight object that will wrap the user's enumerator.
Public Function CreateEnumerator(ByVal obj As IItemData) As stdole.IUnknown
    Dim This As Long
    Dim Struct As UserEnumWrapperType
    
    If mpVTable = 0 Then Init
    
    ' allocate memory to place the new object.
    This = CoTaskMemAlloc(Len(Struct))
    If This = 0 Then Err.Raise 7   ' out of memory
    
    ' tell the user enumerator to reset
    obj.ClassDataLong(eData_EnumReset, vbDefault) = vbDefault
    ' ^^ modified by LaVolpe to use existing Interface function
    
    ' fill the structure of the new wrapper object
    With Struct
        Set .UserEnum = obj
        .cRefs = 1
        .pVTable = mpVTable
    End With
    
    ' move the structure to the allocated memory to complete the object
    CopyMemory ByVal This, ByVal VarPtr(Struct), LenB(Struct)
    ZeroMemory ByVal VarPtr(Struct), LenB(Struct)
    
    ' assign the return value to the newly create object.
    CopyMemory CreateEnumerator, This, 4
End Function

' setup the guids and vtable function pointers.
Private Sub Init()
    InitGUIDS
    InitVTable
End Sub

Private Sub InitGUIDS()
    With IID_IEnumVariant
        .Data1 = &H20404
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    With IID_IUnknown
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Sub

Private Sub InitVTable()
    With mVTable
        .Functions(0) = FuncAddr(AddressOf QueryInterface)
        .Functions(1) = FuncAddr(AddressOf AddRef)
        .Functions(2) = FuncAddr(AddressOf Release)
        .Functions(3) = FuncAddr(AddressOf IEnumVariant_Next)
        .Functions(4) = FuncAddr(AddressOf IEnumVariant_Skip)
        .Functions(5) = FuncAddr(AddressOf IEnumVariant_Reset)
        .Functions(6) = FuncAddr(AddressOf IEnumVariant_Clone)
        
        mpVTable = VarPtr(.Functions(0))
   End With
End Sub

' Helper to get the function pointers of AddressOf methods.
Private Function FuncAddr(ByVal pfn As Long) As Long
    FuncAddr = pfn
End Function


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  VTable functions in the IEnumVariant and IUnknown interfaces.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' When VB queries the interface, we support only two.
' IUnknown
' IEnumVariant
Private Function QueryInterface(ByRef This As UserEnumWrapperType, _
                                ByRef riid As VBGUID, _
                                ByRef pvObj As Long) As Long
    Dim ok As Long
    
    Select Case riid.Data1
        Case IID_IEnumVariant_Data1
            ok = IsEqualGUID(riid, IID_IEnumVariant)
        Case IID_IUnknown_Data1
            ok = IsEqualGUID(riid, IID_IUnknown)
    End Select
    
    If ok Then
        pvObj = VarPtr(This)
        AddRef This
    Else
        QueryInterface = E_NOINTERFACE
    End If
End Function


' increment the number of references to the object.
Private Function AddRef(ByRef This As UserEnumWrapperType) As Long
    With This
        .cRefs = .cRefs + 1
        AddRef = .cRefs
    End With
End Function

' decrement the number of references to the object, checking
' to see if the last reference was released.
Private Function Release(ByRef This As UserEnumWrapperType) As Long
    With This
        .cRefs = .cRefs - 1
        Release = .cRefs
        If .cRefs = 0 Then Delete This
    End With
End Function

' cleans up the lightweight objects and releases the memory
Private Sub Delete(ByRef This As UserEnumWrapperType)
   Set This.UserEnum = Nothing
   CoTaskMemFree VarPtr(This)
End Sub

' move to the next element and return it, signaling if we have reached the end.
Private Function IEnumVariant_Next(ByRef This As UserEnumWrapperType, ByVal celt As Long, ByRef prgVar As Variant, ByVal pceltFetched As Long) As Long
    If This.UserEnum.ClassDataLong(eData_EnumNext, vbDefault) Then
    ' ^^ modified by LaVolpe to use existing Interface function
        VariantCopy prgVar, This.UserEnum.ClassDataObject(eData_EnumItem, vbDefault)
        ' ^^ modified by LaVolpe to use existing Interface function
         
        ' check to see if the pointer is valid (not zero)
        ' before we write to that memory location.
        If pceltFetched Then
            CopyMemory ByVal pceltFetched, 1&, 4
        End If
    Else
        IEnumVariant_Next = ENUM_FINISHED
    End If
End Function

' skip the requested number of elements as long as we don't run out of them.
Private Function IEnumVariant_Skip(ByRef This As UserEnumWrapperType, ByVal celt As Long) As Long
    Do While celt > 0
        If This.UserEnum.ClassDataLong(eData_EnumNext, vbDefault) = vbDefault Then
        ' ^^ modified by LaVolpe to use existing Interface function
            IEnumVariant_Skip = ENUM_FINISHED
            Exit Function
        End If
        celt = celt - 1
    Loop
End Function

' request the user enum to reset.
Private Function IEnumVariant_Reset(ByRef This As UserEnumWrapperType) As Long
   This.UserEnum.ClassDataLong(eData_EnumReset, vbDefault) = vbDefault
   ' ^^ modified by LaVolpe to use existing Interface function
End Function


' we just return a reference to the original object.
Private Function IEnumVariant_Clone(ByRef This As UserEnumWrapperType, ByRef ppenum As stdole.IUnknown) As Long
    CopyMemory ppenum, VarPtr(This), 4
    AddRef This
End Function
