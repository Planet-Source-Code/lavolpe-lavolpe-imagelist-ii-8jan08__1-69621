VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' THE IMAGELIST CLASS (cImageList)
' Child class:  none
' - Returns generic information about the image
' - Stores no data except reference to parent

' Quick List of methods/properties
' ---------------------------------
'     Height returns the height of the image
'     Index gets/sets the index of this image
'     Key gets/sets the key of this image
'     Tag gets/sets the key for this image
'     Width returns the width of the image

Private mParent As IItemData

' Read-Write properties
Public Property Get Key() As String
    Key = mParent.ClassDataString(eData_Key, ObjPtr(Me))
End Property
Public Property Let Key(newKey As String)
    mParent.ClassDataString(eData_Key, ObjPtr(Me)) = newKey
End Property

Public Property Get Index() As Long
    Index = mParent.ClassDataLong(eData_Index, ObjPtr(Me))
End Property
Public Property Let Index(newIndex As Long)
    mParent.ClassDataLong(eData_Index, ObjPtr(Me)) = newIndex
End Property

Public Property Get Tag() As String
    Tag = mParent.ClassDataString(eData_Tag, ObjPtr(Me))
End Property
Public Property Let Tag(newTag As String)
    mParent.ClassDataString(eData_Tag, ObjPtr(Me)) = newTag
End Property

' Read-Only properties
Public Property Get Width() As Long
    Width = mParent.ClassDataLong(eData_Width, ObjPtr(Me))
End Property
Public Property Get Height() As Long
    Height = mParent.ClassDataLong(eData_Height, ObjPtr(Me))
End Property

' Internal use only
Friend Sub xcom_InitializeItem(pIIData As IItemData)
    Set mParent = pIIData
End Sub

Public Function xcom_IEnum() As IUnknown
Attribute xcom_IEnum.VB_UserMemId = -4
Attribute xcom_IEnum.VB_MemberFlags = "40"
    ' IMPORTANT: This property must be set to Procedure ID: -4, manually entered
    ' Also, check the Hide This Member box
    ' See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
    Set xcom_IEnum = Nothing    ' prevents error if user tries to enum this class
End Function

