VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' THE IMAGELIST CLASS (cImageList)
' Child class:  cListImages
'               ^^ child class cListImage
' - Exposes the ListImages classe
' - Returns generic information about the image list
' See below for more

' Quick List of methods/properties
' ---------------------------------
'     Height returns the height for every image in the list
'     Images returns the cListImages class for this image list
'     Index gets/sets the index of this image list
'     Key gets/sets the key of this image list
'     MaxImages returns maximum number of images that are allowed on this image list
'     Width returns the width for every image in the list


Private mParent As IItemData
Private m_Width As Long
Private m_Height As Long
Private m_Images As cListImages

Implements IItemData

Public Property Get Images() As cListImages
    ' returns the cListImages class
    Set Images = m_Images
End Property

' returns/set the key for this image list
Public Property Get Key() As String
    Key = mParent.ClassDataString(eData_Key, ObjPtr(Me))
End Property
Public Property Let Key(newKey As String)
    mParent.ClassDataString(eData_Key, ObjPtr(Me)) = newKey
End Property

' returns/set the index for this image list
Public Property Let Index(newIndex As Long)
    mParent.ClassDataLong(eData_Index, ObjPtr(Me)) = newIndex
End Property
Public Property Get Index() As Long
    Index = mParent.ClassDataLong(eData_Index, ObjPtr(Me))
End Property

' returns the width/height of images within this image list
Public Property Get Width() As Long
    Width = m_Width
End Property
Public Property Get Height() As Long
    Height = m_Height
End Property

Public Property Get MaxImages() As Long
    MaxImages = mParent.ClassDataLong(eData_MaxImgCount, ObjPtr(Me))
End Property

' Following property allows child classes to set/get this class's properties
Private Property Let IItemData_ClassDataLong(DataName As eDataNames, oID As Long, ByVal newValue As Long)
    Select Case DataName
        Case eData_Width: m_Width = newValue
        Case eData_Height: m_Height = newValue
    End Select
End Property
Private Property Get IItemData_ClassDataLong(DataName As eDataNames, oID As Long) As Long
    Select Case DataName
        Case eData_Width: IItemData_ClassDataLong = m_Width
        Case eData_Height: IItemData_ClassDataLong = m_Height
        Case eData_GDItoken: IItemData_ClassDataLong = mParent.ClassDataLong(DataName, oID)
        Case eData_MaxImgCount: IItemData_ClassDataLong = mParent.ClassDataLong(DataName, ObjPtr(Me))
    End Select
End Property

Private Property Get IItemData_ClassDataObject(DataName As eDataNames, oID As Long) As Variant
    ' not used for this class
End Property

Private Property Let IItemData_ClassDataString(DataName As eDataNames, oID As Long, ByVal newValue As String)
    ' not used, this class has no string values to share
End Property
Private Property Get IItemData_ClassDataString(DataName As eDataNames, oID As Long) As String
    ' not used, this class has no string values to share
End Property


Friend Sub xcom_InitializeList(Width As Long, Height As Long, initCount As Long, pIIData As IItemData)
    ' internal use only. The cImageLists class calls this to pass initialization values
    m_Width = Width
    m_Height = Height
    Set mParent = pIIData
    Set m_Images = New cListImages
    Call m_Images.xcom_InitializeImages(Me, Width, Height, initCount)
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

