VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' This class is used as an interface between parent & child classes so the child
' can get information from the parent or set the parent's properties if needed
' It has also been modified to act as an IEnum interface for the
'   cImageLists & cListImages classes

' The call back routines are "private" so they are never exposed to the user

' Additional info.
' The cListImage class is a dummy class. It stores nothing
'       except a reference back to it its parent class of cListImages
'   Its only purpose is to provide the user with an image object they
'   can query properties from, like Index, Key, Width, Height, Tag

' The cImageList class is nearly a dummy class. It stores only the
'       width & height of each image in the image list and also
'       stores a reference back to its parent class of cImageLists
'   Its only purpose is to expose the cListImages class and provide
'   some information to the user like Index, Key, Width, Height

' This backward communication allows us to add a Key and Index property
' to the cListImage and cImageList classes which normal collections do not
' allow. It also allows us to change it at that level too, adding a bit more
' intuitive coding for you.
' For example to change a key: .ImageLists(x).Images(n).Key = "newKey"
' and the Images(n) calls back to parent and says, "My key changed to newKey".

' A little extra work but worth it in my opinion

Public Enum eDataNames
    eData_Key = 0       ' return/set child's key
    eData_Index = 1     ' return/set child's index
    eData_Width = 2     ' return child's width or set parent's width
    eData_Height = 3    ' return child's height or set parent's height
    eData_Tag = 4       ' return/set child's tag
    eData_GDItoken = 5  ' return parent's GDI token, if any
    eData_MaxImgCount = 6 ' return number of images that will fit on a particular image list
    ' IENUM interface portion
    eData_EnumReset = 7 ' tell IEnum to reset, ready to enum first item
    eData_EnumNext = 8  ' ask IEnum if there are more items to enum
    eData_EnumItem = 9  ' return specific item to IEnum
End Enum

Public Property Get ClassDataLong(DataName As eDataNames, oID As Long) As Long
Attribute ClassDataLong.VB_MemberFlags = "40"
    ' returns a property value as Long
End Property
Public Property Let ClassDataLong(DataName As eDataNames, oID As Long, ByVal newValue As Long)
    ' sets a Long property value
End Property

Public Property Get ClassDataString(DataName As eDataNames, oID As Long) As String
Attribute ClassDataString.VB_MemberFlags = "40"
    ' returns a property value as String
End Property
Public Property Let ClassDataString(DataName As eDataNames, oID As Long, ByVal newValue As String)
    ' sets a String property value
End Property

Public Property Get ClassDataObject(DataName As eDataNames, oID As Long) As Variant
    ' used for IEnum only
End Property
