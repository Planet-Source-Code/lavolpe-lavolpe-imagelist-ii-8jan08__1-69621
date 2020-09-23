VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListImages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' THE IMAGES CLASS (cListImages)
' Child class: cListImage
' - Manages collection of images, their keys and tags
' - Creates and destroys images
' - Returns specific image classes
' - Renders images and exports images as icons
' Also see below


' Quick List of methods/properties
' ---------------------------------
'     AddImage_fromClipBoard loads image from clipboard
'     AddImage_fromFile loads image from file
'     AddImage_fromHandle loads image from memory handle
'     AddImage_fromResource loads image from compiled resource file
'     AddImage_fromStdPicture loads image from standard picture object
'     AddImage_fromStream loads image from array/stream
'     Clear destroys all images in the list
'     Count returns number of images
'     DelayLoaded gets/sets whether images are created immediately or not (see property's remarks)
'     ExtractIcon converts image to icon
'     ExtracIconEx converts image to icon after custom-processing the image
'     IsKeyAssigned returns whether or not a passed Key is unique
'     Item returns a specific cListImage class
'     MaxImages returns maximum number of images that are allowed on this image list
'     Remove deletes a specific image
'     Render draws an image
'     RenderRange draws a range of images as a single strip

' Notes about indexing
' 1. Referencing by Ordinal index is always faster than referencing by Key.
' 2. The Keys collection is sorted and finding a key is very fast using binary searching.
'   The key has a reference to the cListImage class pointer so we can obtain a specific
'   class from the key very fast.
' 3. Likewise, the Pointers collection is sorted and used when a child class is calling
'   back to this class. The pointer will be used to identify the calling child class
' 4. ReIndexing pointers and keys happens when an image is inserted before the end of
'   the image list, when changing the Index property of an image, or deleting an image.
'   Actually, keys are never reindexed because they only use pointer references which
'   don't change. But the Pointer collection must be reindexed because it also contains
'   references to the image's ordinal index in the image list. This reindexing will use a
'   combination of binary searches to reindex or a complete collection iteration, depending
'   on the amount of images that need to be reindexed.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Function CreateIconFromResourceEx Lib "user32.dll" (ByRef presbits As Any, ByVal dwResSize As Long, ByVal fIcon As Long, ByVal dwVer As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal Flags As Long) As Long

Private Enum eCloneActions
    cloneCreate = 1
    cloneRestore = 2
    cloneRelease = 3
End Enum

Private Type SafeArrayBound
    cElements As Long
    lLbound As Long
End Type
Private Type SafeArray        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SafeArrayBound ' 32 bytes as used. Can be used for 1D and/or 2D arrays
End Type

Private Type ImageDataStructure  '8 bytes
    ImageClass As cListImage
    Key As Long
End Type
Private Type KeyDataStructure   ' 8 bytes
    Key As String
    ImagePtr As Long
End Type
Private Type PointerStructure   ' 8 bytes
    Index As Long
    ImagePtr As Long
End Type

Private mParent As IItemData                ' reference to parent class: cImageList
Private m_Clone As cListImages              ' property page use only; allows undo of changes

Private m_Images() As ImageDataStructure    ' images to key cross-reference
Private m_Keys() As KeyDataStructure        ' keys to objptr cross-reference
Private m_Indexes() As PointerStructure     ' objptr to index cross-reference
Private m_Tags() As String                  ' image tags

Private m_ItemCount As Long                 ' number of images & pointers
Private m_KeyCount As Long                  ' number of keys can be less than m_Count
Private m_Strip As c32bppDIB                ' the imagelist strip
Private m_Slot As c32bppDIB
Private m_Width As Long                     ' list width, cached in this class 'cause of frequent use
Private m_Height As Long                    ' list height, cached in this class 'cause of frequent use
Private m_DelayLoad As Boolean              ' delay load option. See Me.DelayLoaded
Private m_EnumIndex As Long
Implements IItemData                        ' allows child cListImage class to call back to this class

Public Property Get Item(Index As Variant) As cListImage
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = local_KeyToObject(Index)   ' returns a specific image
    ' IMPORTANT: This property must be set to Procedure ID: Default and also
    ' checked as User Interface Default. See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
End Property

' Various ways of adding an image to the imagelist
Public Function AddImage_fromFile(ByVal FileName As String, _
                                Optional ByVal Index As Variant, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal iconBitDepth As Long = 32, _
                                Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   FileName :: full path & file (unicode supported) containing an image
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_File(FileName, m_Width, m_Height, False, iconBitDepth) Then
        Set AddImage_fromFile = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromStream(inStream() As Byte, _
                                    Optional ByVal Index As Variant, _
                                    Optional ByVal Key As String = vbNullString, _
                                    Optional ByVal Tag As String = vbNullString, _
                                    Optional ByVal streamStart As Long = 0&, _
                                    Optional ByVal streamLength As Long = 0&, _
                                    Optional ByVal iconBitDepth As Long = 32, _
                                    Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   inStream :: a single dimensioned array containing a complete image
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   streamStart :: byte position in array where image begins
    '   streamLength :: number of bytes that comprise the image
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_Stream(inStream, m_Width, m_Height, streamStart, streamLength, False, iconBitDepth) Then
        Set AddImage_fromStream = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromResource(ByVal ResIndex As Variant, ByVal ResSection As Variant, _
                            VBglobal As IUnknown, _
                            Optional ByVal Index As Variant, _
                            Optional ByVal Key As String = vbNullString, _
                            Optional ByVal Tag As String = vbNullString, _
                            Optional ByVal streamStart As Long = 0&, _
                            Optional ByVal streamLength As Long = 0&, _
                            Optional ByVal iconBitDepth As Long = 32, _
                            Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   ResIndex :: the resource index containing the image
    '   ResSection :: the resource section (i.e., "Custom", vbResIcon)
    '   VBglobal :: always pass as VB.Global
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   streamStart :: byte position in resource array where image begins (usually LBound)
    '   streamLength :: number of bytes that comprise the image (usually UBound+1)
    '   iconBitDepth :: when file is an icon, desired resolution
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    If TypeOf VBglobal Is VB.Global Then
        Dim newImg As c32bppDIB
        Set newImg = New c32bppDIB
        If newImg.LoadPicture_Resource(ResIndex, ResSection, VBglobal, m_Width, m_Height, streamStart, streamLength, iconBitDepth) Then
            Set AddImage_fromResource = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
        End If
    End If
End Function
Public Function AddImage_fromStdPicture(Picture As StdPicture, _
                                Optional ByVal Index As Variant, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   Picture :: a VB standard picture object
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_StdPicture(Picture) Then
        Set AddImage_fromStdPicture = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromHandle(Handle As Long, _
                                Optional ByVal Index As Variant, _
                                Optional ByVal Key As String = vbNullString, _
                                Optional ByVal Tag As String = vbNullString, _
                                Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   Handle :: a memory handle to the image
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_ByHandle(Handle) Then
        Set AddImage_fromHandle = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
    End If
End Function
Public Function AddImage_fromClipBoard(Optional ByVal Index As Variant, _
                                    Optional ByVal Key As String = vbNullString, _
                                    Optional ByVal Tag As String = vbNullString, _
                                    Optional ByVal Scaled As Boolean = True) As cListImage
    ' Parameters:
    '   Index :: placement in the image list. Default is at the end
    '   Key :: key for the loaded image
    '   Tag :: tag for the loaded image
    '   Scaled :: False if image is stretched, otherwise image is scaled proportionally
    
    Dim newImg As c32bppDIB
    Set newImg = New c32bppDIB
    If newImg.LoadPicture_ClipBoard() Then
        Set AddImage_fromClipBoard = local_AddItem(newImg, Index, Key, Tag, Scaled, 1&)
    End If
End Function

' Render function is a wrapper around the c32bppDIB class
' --------------------------------------------------------
Public Function Render(Index As Variant, ByVal destinationDC As Long, _
                Optional ByVal destX As Long, Optional ByVal destY As Long, _
                Optional ByVal destWidth As Long, Optional ByVal destHeight As Long, _
                Optional ByVal SrcX As Long, Optional ByVal SrcY As Long, _
                Optional ByVal srcWidth As Long, Optional ByVal srcHeight As Long, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal grayScale As eILGrayScales = lvil_gsclNone, _
                Optional ByVal LightAdjustment As Single = 0!, _
                Optional ByVal Angle As Single = 0!, _
                Optional ByVal CenterOnDestXY As Boolean = False) As Boolean
                
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be rendered
    '   destinationDC :: target DC to render image to
    '   destX,destY :: target DC coordinates to be rendered at
    '   destWidth,destHeight :: dimensions to be rendered to
    '   SrcX,SrcY :: image coordinates to being rendering from
    '   srcWdith,srcHeight :: image dimensions that will be rendered from SrcX,Srcy to destX,destY
    '   Opacity :: value btwn 0 & 100. 100 is fully opaque, 0 is fully transparent
    '   grayScale :: one of several gray scale formulas
    '   LigthAdjustment :: values btwn -100 to 100 which is a percentage of lightness(+) or darkness(-) to add to image
    '   Angle :: values btwn -360 and 360 to rotate image. Rotation is clockwise use negative values for counterclockwise rotation
    '   CenterOnDestXY :: if True, image is rendered around destX,destY else image is rendered from destX,destY

    Dim iIndex As Long
    If m_ItemCount = vbDefault Then Exit Function
    Call local_KeyToObject(Index, iIndex)
    If iIndex = vbDefault Then
        Err.Raise 1962, "Images.Render", "Invalid Key or Index"
        Exit Function
    End If
    
    If m_Strip.ImageType = imgError Then
        If local_AdjustDibSize(vbDefault, vbDefault, vbDefault) = False Then Exit Function
    End If
    
    If m_Slot Is Nothing Then
        Set m_Slot = New c32bppDIB
        m_Slot.InitializeDIB m_Width, m_Height
    End If
    
    Dim oSA As SafeArray, sSA As SafeArray
    Dim aSrc() As Long, aDst() As Long
    Dim Y As Long, X As Long, lScanWidth As Long
    
    ' validate optional parameters for source image, translating Index to slot in image list
    ' Regarding mirroring. If the source or destination width or height is negative, mirroring will be applied
    If destWidth = 0& Then destWidth = m_Width
    If destHeight = 0& Then destHeight = m_Height
    If srcWidth = 0& Then
        srcWidth = m_Width
    ElseIf srcWidth > m_Width Then
        srcWidth = m_Width
    End If
    If srcWidth < 0& Then           ' locally, only the destWidth can be negative for mirroring, not srcHeight
        destWidth = -destWidth
        srcWidth = -srcWidth
    End If
    If srcHeight = 0& Then
        srcHeight = m_Height
    ElseIf srcHeight > m_Height Then
        srcHeight = m_Height
    End If
    If srcHeight < 0& Then          ' locally, only the destHeight can be negative for mirroring, not srcHeight
        destHeight = -destHeight
        srcHeight = -srcHeight
    End If
    If SrcX < 0& Then SrcX = 0&
    If SrcY < 0& Then SrcY = 0&
    If SrcX + srcWidth > m_Width Then srcWidth = m_Width - SrcX
    If SrcY + srcHeight > m_Height Then srcHeight = m_Height - SrcY
    
    iparseOverlayHost_Long aSrc(), VarPtr(oSA), 2, m_Height, m_Strip.Width, m_Strip.BitsPointer
    iparseOverlayHost_Long aDst(), VarPtr(sSA), 2, m_Height, m_Width, m_Slot.BitsPointer
    X = (iIndex - 1) * m_Width
    lScanWidth = m_Slot.scanWidth
    For Y = 0 To m_Height - 1
        CopyMemory aDst(0, Y), aSrc(X, Y), lScanWidth
    Next
    iparseOverlayHost_Long aSrc(), 0, 0, 0, 0, 0
    iparseOverlayHost_Long aDst(), 0, 0, 0, 0, 0
    Render = m_Slot.Render(destinationDC, destX, destY, destWidth, destHeight, SrcX, SrcY, srcWidth, srcHeight, Opacity, , , , grayScale, LightAdjustment, Angle, CenterOnDestXY)

End Function

Public Sub RenderRange(ByVal StartIndex As Variant, ByVal StopIndex As Variant, destinationDC As Long, _
                        ByVal destX As Long, ByVal destY As Long, _
                        ByVal imageWidth As Long, ByVal imageHeight As Long)
    
    ' Allows rendering a range of images as a strip. Useful if providing a run-time GUI imagelist for users
    
    ' Parameters
    '   StartIndex :: Ordinal or key that begins the range.
    '   StopIndex :: Ordinal or key that ends the range. If exceeds nr images, last image is used
    '   destinationDC :: the DC to render the range to
    '   destX,destY :: the Left & Top coordinate to begin rendering at
    '   imageWidth :: the width each image should be rendered as
    '   imageHeight :: the heigh each image should be rendered as
    
    If m_ItemCount = vbDefault Then Exit Sub
    
    On Error Resume Next
    Dim Index As Long, bUnCompiled As Boolean
    Dim xOffset As Long, yOffset As Long
    Dim startItem As Long, stopItem As Long
    
    If local_KeyToObject(StartIndex, startItem) Is Nothing Then Exit Sub
    If local_KeyToObject(StopIndex, stopItem) Is Nothing Then
        stopItem = m_ItemCount
    ElseIf stopItem < startItem Then
        Exit Sub
    End If
    
    If m_Strip.ImageType = imgError Then
        If local_AdjustDibSize(vbDefault, vbDefault, vbDefault) = False Then Exit Sub
    End If
    
    ' do some efficiency enhancements while rendering
    If m_Strip.isGDIplusEnabled Then
        Debug.Assert local_IsUnCompiled(bUnCompiled)
        If bUnCompiled Then m_Strip.KeepGDIplusActive = True ' prevent reloading GDI+ image for every subimage
    Else
        m_Strip.ManageOwnDC = True  ' prevent creating DC for each sub-image drawn, create once
    End If
    
    startItem = startItem - 1
    If m_Width < imageWidth Then
        xOffset = (imageWidth - m_Width) \ 2
        yOffset = (imageHeight - m_Height) \ 2
        For Index = startItem To stopItem - 1
            m_Strip.Render destinationDC, destX + xOffset, destY + yOffset, m_Width, m_Height, Index * m_Width, vbDefault, m_Width, m_Height
            destX = destX + imageWidth
        Next
    Else
        For Index = startItem To stopItem - 1
            m_Strip.Render destinationDC, destX, destY, imageWidth, imageHeight, Index * m_Width, 0&, m_Width, m_Height
            destX = destX + imageWidth
        Next
    End If
    
    ' release efficiency enhancements as needed
    If m_Strip.isGDIplusEnabled Then
        If bUnCompiled Then m_Strip.KeepGDIplusActive = False ' can crash IDE if END executed
    Else
        m_Strip.ManageOwnDC = False ' destroy DC as needed
    End If
    
End Sub

Public Function ExtractIcon(Index As Variant, Optional Rtn_BitDepth As Long) As Long

    ' Function extracts an image and converts it to an icon
    ' Return value of the function is the icon handle if it was created.
    ' Icons must be destroyed by you with a call to DestroyIcon API
    
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be iconized
    
    ' Notes:
    
    ' 1. Depending on the per-pixel translucency and number of colors
    ' within the image, the returned icon can exist in one of several
    ' bit depths: 1,4,8 bit paletted, 24 bit true color, or 32 bit alpha-blended
    
    ' 2. When high rotation and/or quality interpolation is used and an image is stretched
    ' or otherwise resized to fit the image list, that image will probably be
    ' alphablended along its edges even if the image previously was 100% opaque.
    ' Therefore, resized images very well may return a 32bpp icon vs 24bpp or lower

    If m_ItemCount = 0& Then Exit Function
    If m_Strip.ImageType = imgError Then
        ' when DelayLoad is active, no cImage class will be created yet
        If local_AdjustDibSize(vbDefault, vbDefault, vbDefault) = False Then Exit Function
    End If
    Dim iIndex As Long
    Call local_KeyToObject(Index, iIndex)
    If iIndex = 0& Then
        Err.Raise 1962, "Images.ExtractIcon", "Invalid Key or Index"
    Else
        ExtractIcon = local_ExtractIcon(m_Strip, iIndex, Rtn_BitDepth, m_Width, m_Height)
    End If
End Function

Public Function ExtractIconEx(Index As Variant, _
                Optional ByVal MirrorAxisX As Boolean = False, _
                Optional ByVal MirrorAxisY As Boolean = False, _
                Optional ByVal Opacity As Long = 100&, _
                Optional ByVal grayScale As eILGrayScales = lvil_gsclNone, _
                Optional ByVal LightAdjustment As Single = 0!, _
                Optional ByVal Angle As Single = 0!, _
                Optional Rtn_iconWidth As Long, Optional Rtn_iconHeight As Long, Optional Rtn_BitDepth As Long) As Long
                
    
    ' Function will return an icon in a rendered appearance other than the original image
    ' See notes in ExtractIcon also
                
    ' Parameters:
    '   Index :: Ordinal or Key for the image to be rendered
    '   MirroredAxisX :: Set to true if image is to be mirrored horizontally
    '   MirroredAxisY :: Set to true if image is to be mirrored vertically
    '   Opacity :: value btwn 0 & 100. 100 is fully opaque, 0 is fully transparent
    '   grayScale :: one of several gray scale formulas
    '   LigthAdjustment :: values btwn -100 to 100 which is a percentage of lightness(+) or darkness(-) to add to image
    '   Angle :: values btwn -360 and 360 to rotate image. Rotation is clockwise use negative values for counterclockwise rotation
    '   ... Rotation will create a larger icon than the source with the source image rotated around the center of the icon
    '   Rtn_iconWidth [out]:: will be the width of the created icon.
    '   Rtn_iconHeight [out]:: will be the height of the created icon
    '   Rtn_BitDepth [out]:: will be the bit depth of the icon: 1,4,8,24,32

    If m_ItemCount = 0& Then Exit Function
    If m_Strip.ImageType = imgError Then
        ' when DelayLoad is active, no cImage class will be created yet
        If local_AdjustDibSize(vbDefault, vbDefault, vbDefault) = False Then Exit Function
    End If
    
    Dim iIndex As Long
    Call local_KeyToObject(Index, iIndex)
    If iIndex = 0& Then
        Err.Raise 1962, "Images.ExtractIconEx", "Invalid Key or Index"
        Exit Function
    Else
        ' validate optional parameters for source image, translating Index to slot in image list
        ' Regarding mirroring. If the source or destination width or height is negative, mirroring will be applied
        Dim cX As Long, cY As Long, X As Long
        Dim destX As Long, destY As Long
        Dim tImage As c32bppDIB
        If MirrorAxisY Then cY = -m_Height Else cY = m_Height
        If MirrorAxisX Then
            cX = -m_Width
            ' Mirroring: when using GDI+, we have to tweak the offset.
            ' All images are one bitmap. When GDI+ mirrors the entire bitmap then the
            ' index will be pointing to the wrong subimage. Fix it here.
            If m_Strip.isGDIplusEnabled Then iIndex = UBound(m_Images) - iIndex + 1
            ' Note: when GDI+ not available, the manual routines do this naturally
        Else
            cX = m_Width
        End If
        X = m_Width * (iIndex - 1&)

        Set tImage = New c32bppDIB
        
        ' reduce icon size if possible. Can be done when rotation at 90 degree increments
        If Angle = Int(Angle) Then                      ' else degree fraction in use
            If Angle Mod 360 = 0 Then                   ' no rotation?
                tImage.InitializeDIB m_Width, m_Height
                Angle = 0!
            ElseIf Angle Mod 180 = 0 Then                   ' vertical rotation only?
                tImage.InitializeDIB m_Width, m_Height
                destY = m_Height \ 2: destX = m_Width \ 2
            ElseIf Angle Mod 90 = 0 Then                    ' 90 degree rotation?
                tImage.InitializeDIB m_Height, m_Width
                destY = m_Width \ 2: destX = m_Height \ 2   ' calc center point
            End If
        End If
        If tImage.ImageType = imgError Then             ' angle did not fit one of above categories
            destX = Sqr(m_Width * m_Width + m_Height * m_Height)    ' create icon to allow any angle rotation
            tImage.InitializeDIB destX, destX                       ' create the temp image
            destY = destX \ 2: destX = destY                        ' calc center point
        End If
        tImage.gdiToken = m_Strip.gdiToken
            
        ' render image to a temporary dib
        m_Strip.Render 0&, destX, destY, m_Width, m_Height, X, 0, cX, m_Height, Opacity, , , tImage, grayScale, LightAdjustment, Angle, (Angle <> 0!)
        Rtn_iconWidth = tImage.Width
        Rtn_iconHeight = tImage.Height
        ' call local function to return icon from that dib
        ExtractIconEx = local_ExtractIcon(tImage, 1, Rtn_BitDepth, tImage.Width, tImage.Height)
        Set tImage = Nothing
    End If
End Function

' property returns/sets whether delay load option is applied
' The delay load option can help conserve resources when multiple
' imagelists exist but may not be utilized. When set, the property
' prevents creating the actual image list (DIB) until the image
' is required: Adding, Removing, Rendering or Indexing images
Public Property Let DelayLoaded(doDelay As Boolean)
    m_DelayLoad = doDelay
End Property
Public Property Get DelayLoaded() As Boolean
    DelayLoaded = m_DelayLoad
End Property

Public Property Get Count() As Long ' return number of images managed
    Count = m_ItemCount
End Property

Public Sub Clear() ' clear all images
    Dim bUnCompiled As Boolean
    Erase m_Images
    Erase m_Keys
    Erase m_Indexes
    m_ItemCount = vbDefault
    m_KeyCount = vbDefault
    If Not m_Strip Is Nothing Then m_Strip.DestroyDIB
End Sub

Public Sub Remove(Index As Variant) ' deletes a specific image

    If m_ItemCount = vbDefault Then Exit Sub
    
    Dim pIndex As Long, iIndex As Long, kIndex As Long
    Dim sKey As String
    Dim tImage As cListImage
    
    ' validate index
    Call local_KeyToObject(Index, iIndex)
    If Not iIndex = vbDefault Then
        ' find the index from the pointer
        If m_ItemCount = 1& Then
            Me.Clear
        Else
            
            pIndex = local_FindPointer(ObjPtr(m_Images(iIndex).ImageClass), True, m_ItemCount)
            
            If local_AdjustDibSize(m_ItemCount, iIndex, -1&) = False Then Exit Sub
            ' ^^ can't think of any reason why this could happen
            
            If m_Images(iIndex).Key Then                ' does it have a key?
                If m_KeyCount = 1& Then
                    m_KeyCount = vbDefault
                    Erase m_Keys
                Else
                    ' get the key from its pointer
                    CopyMemory ByVal VarPtr(sKey), m_Images(iIndex).Key, 4&
                    kIndex = local_FindKey(sKey, True)      ' find the key's index
                    ' clear the StrPtr in our m_Keys collection
                    CopyMemory ByVal VarPtr(m_Keys(kIndex).Key), vbDefault, 4&
                    sKey = vbNullString                     ' clear the key
                    ' shift the keys collection
                    If kIndex < m_KeyCount Then
                        CopyMemory ByVal VarPtr(m_Keys(kIndex)), ByVal VarPtr(m_Keys(kIndex + 1)), (m_KeyCount - kIndex) * 8&
                        CopyMemory ByVal VarPtr(m_Keys(m_KeyCount).Key), vbDefault, 4&
                    End If
                    m_KeyCount = m_KeyCount - 1
                    If m_KeyCount + 3 < UBound(m_Keys) Then ReDim Preserve m_Keys(1 To m_KeyCount)
                End If
            End If
            local_ReIndexPointers iIndex + 1, m_ItemCount, -1, False
            Set m_Images(iIndex).ImageClass = Nothing   ' clear the class & its images
            m_Tags(iIndex) = vbNullString         ' clear its tag
            
            If iIndex < m_ItemCount Then                ' need to shift some arrays
                ' shift the images collection
                CopyMemory ByVal VarPtr(m_Images(iIndex)), ByVal VarPtr(m_Images(iIndex + 1)), (m_ItemCount - iIndex) * 8&
                CopyMemory ByVal VarPtr(m_Tags(iIndex)), ByVal VarPtr(m_Tags(iIndex + 1)), (m_ItemCount - iIndex) * 4&
                CopyMemory ByVal VarPtr(m_Tags(m_ItemCount)), vbDefault, 4&
                FillMemory ByVal VarPtr(m_Images(m_ItemCount)), 8&, 0
            End If
            If pIndex < m_ItemCount Then
                ' shift the pointers collection
                CopyMemory ByVal VarPtr(m_Indexes(pIndex)), ByVal VarPtr(m_Indexes(pIndex + 1)), (m_ItemCount - pIndex) * 8&
            End If
            
            m_ItemCount = m_ItemCount - 1
            If m_ItemCount + 3 < UBound(m_Images) Then
                ReDim Preserve m_Images(1 To m_ItemCount)
                ReDim Preserve m_Indexes(1 To m_ItemCount)
                ReDim Preserve m_Tags(1 To m_ItemCount)
            End If
            
        End If
    End If
End Sub

Public Function IsKeyAssigned(Key As String) As Long
    ' Returns whether or not passed Key has been assigned
    ' Return value of 0=unique key, else is the Index of the image having that key
    Dim kIndex As Long, bUnique As Boolean
    kIndex = local_FindKey(Key, bUnique)
    If Not bUnique Then IsKeyAssigned = local_PointerToIndex(m_Keys(kIndex).ImagePtr)
End Function

Public Property Get MaxImages() As Long
    MaxImages = mParent.ClassDataLong(eData_MaxImgCount, vbDefault)
End Property


' LOCAL USE ONLY ROUTINES

Private Function local_AddItem(cDIB As c32bppDIB, InsertIndex As Variant, _
                        Key As String, Tag As String, Scaled As Boolean, ByVal SlotsToShift As Long) As cListImage
    
    ' Local function to append an image at a specific location. Called from the various AddImage functions & xcom_AddItem_FromDIB
    ' Parameters:
    '   cDIB :: a loaded image in 32bpp format, scaled or stretched appropriately
    '   InsertIndex :: ordinal or key of image to insert new image at. Values < 0 or > image count defaults to end of list
    '   Key :: the key for the new image
    '   Tag :: the tag for the new image
    '   SlotsToShift :: number of blank images we will be inserting, default is one
    '       When value is zero, a bulk load is being performed from the propertypage
    '       and is only called by xcom_AddItem_FromDIB. In this case, that function already resized the DIB and inserted
    '       blank placeholders into the image list as needed
    
    Dim newImage As cListImage
    Dim iIndex As Long, InsertAt As Long
    Dim bUnique As Boolean, kIndex As Long
    Dim imgCx As Long, imgCy As Long
    Dim imgX As Long, imgY As Long
    
    ' validate imagelist has room
    If (m_ItemCount + 1) > (1024000 \ (m_Width * m_Height)) Then
        Err.Raise 1973, "Images.Add", "Maximum number of images have been exceeded. Create new image list"
        Exit Function
    End If
    
    If Not Key = vbNullString Then
        kIndex = local_FindKey(Key, bUnique)
        If Not bUnique Then
            Err.Raise 1963, "Images.Add", "Duplicated Key. Key already exists in this collection"
            Exit Function
        End If
    End If
    
    ' resize array as needed, leaving a small buffer
    If iparseIsArrayEmpty(VarPtrArray(m_Images)) = vbDefault Then
        ReDim m_Images(1 To 3)
        ReDim m_Indexes(1 To 3)
        ReDim m_Tags(1 To 3)
    ElseIf m_ItemCount = UBound(m_Images) Then
        ReDim Preserve m_Images(1 To m_ItemCount + 3)
        ReDim Preserve m_Indexes(1 To m_ItemCount + 3)
        ReDim Preserve m_Tags(1 To m_ItemCount + 3)
    End If
    
    Set newImage = local_KeyToObject(InsertIndex)       ' find the insert position
    m_ItemCount = m_ItemCount + 1                       ' increment image count
    If newImage Is Nothing Then                         ' invalid position; use end of list
       InsertAt = m_ItemCount
    Else
        InsertAt = local_PointerToIndex(ObjPtr(newImage)) ' use its position to insert into
    End If
    Set newImage = New cListImage
    
    If Not SlotsToShift = vbDefault Then                ' resize DIB as needed
        If local_AdjustDibSize(m_ItemCount, InsertAt, SlotsToShift) = False Then Exit Function
    End If
    
    If InsertAt < m_ItemCount Then                      ' inserting before last image
        local_ReIndexPointers InsertAt, m_ItemCount, 1, True  ' re-index items as needed
        ' shift the image & tag collections right one spot
        CopyMemory ByVal VarPtr(m_Images(InsertAt + 1)), ByVal VarPtr(m_Images(InsertAt)), (m_ItemCount - InsertAt) * 8&
        CopyMemory ByVal VarPtr(m_Tags(InsertAt + 1)), ByVal VarPtr(m_Tags(InsertAt)), (m_ItemCount - InsertAt) * 4&
        CopyMemory ByVal VarPtr(m_Images(InsertAt).ImageClass), vbDefault, 4&
        CopyMemory ByVal VarPtr(m_Tags(InsertAt)), vbDefault, 4&
    End If
    
    Set m_Images(InsertAt).ImageClass = newImage        ' set new image to the new position
    m_Images(InsertAt).Key = StrPtr(Key)                ' copy the key reference, if any, to the new position
    CopyMemory ByVal VarPtr(m_Tags(InsertAt)), StrPtr(Tag), 4&    ' move the tag to our new item
    CopyMemory ByVal VarPtr(Tag), vbDefault, 4&
    
    If Not kIndex = vbDefault Then                      ' add the key if needed.
        local_AddKey Key, kIndex, ObjPtr(newImage)
        CopyMemory ByVal VarPtr(Key), vbDefault, 4&
    End If
    
    local_AddPointer ObjPtr(newImage), InsertAt, True   ' add index to the pointer-to-index cross-reference
    
    If Scaled Then                                      ' scale keeping aspect ratio if requested
        cDIB.ScaleImage m_Width, m_Height, imgCx, imgCy, eScaleOptions.ScaleToSize
        imgX = (m_Width - imgCx) \ 2
        imgY = (m_Height - imgCy) \ 2
    Else                                                ' stretching
        imgCx = m_Width
        imgCy = m_Height
    End If
        
    ' we need to add the image to the imagelist now
    Call cDIB.Render(0&, (InsertAt - 1) * m_Width + imgX, imgY, imgCx, imgCy, , , , , , , , m_Strip)
    ' finally, initialize image passing us as its parent & return result
    Call newImage.xcom_InitializeItem(Me)
    Set local_AddItem = newImage

End Function

Private Sub local_ReIndexPointers(ByVal startIdx As Long, ByVal stopIdx As Long, incrMode As Long, bAppending As Boolean)
    
    ' Manages the Pointer to Index array
    ' The cListImage class stores no data, they call back to this class to
    ' retrieve their key, tag & index values by passing thier class pointer
    
    Dim pIndex As Long
    ' binary search is fast for single items, but not so fast if many are needed to be searched
    If (stopIdx - startIdx) < m_ItemCount \ 3 Then  ' use binary search
        For startIdx = startIdx To stopIdx
            pIndex = local_FindPointer(ObjPtr(m_Images(startIdx).ImageClass), True, m_ItemCount + bAppending)
            m_Indexes(pIndex).Index = m_Indexes(pIndex).Index + incrMode
        Next
    Else
        startIdx = startIdx - 1
        stopIdx = stopIdx + 1
        For pIndex = 1 To m_ItemCount               ' iterate the entire collection
            If m_Indexes(pIndex).Index > startIdx Then
                If m_Indexes(pIndex).Index < stopIdx Then m_Indexes(pIndex).Index = m_Indexes(pIndex).Index + incrMode
            End If
        Next
    End If
End Sub

Private Sub local_AddKey(Key As String, ByVal KeyPos As Long, KeyRef As Long)

    ' Adds a new key to the collection

    If KeyPos = vbDefault Then KeyPos = local_FindKey(Key, True)
    
    If m_KeyCount = vbDefault Then
        ReDim m_Keys(1 To 3)
    ElseIf m_KeyCount = UBound(m_Keys) Then
        ReDim Preserve m_Keys(1 To m_KeyCount + 3)
    End If
    m_KeyCount = m_KeyCount + 1
    
    If KeyPos < m_KeyCount Then
        CopyMemory ByVal VarPtr(m_Keys(KeyPos + 1)), ByVal VarPtr(m_Keys(KeyPos)), (m_KeyCount - KeyPos) * 8&
    End If
    ' move the key to new slot
    CopyMemory ByVal VarPtr(m_Keys(KeyPos).Key), StrPtr(Key), 4&
    ' reference the image class
    m_Keys(KeyPos).ImagePtr = KeyRef
    
End Sub

Private Sub local_ReyKeyImage(Index As Long, newKey As String)

    ' Changes a key for a specific item

    Dim kIndex As Long, newKindex As Long
    Dim bUnique As Boolean, tempKey As String
    
    If m_Images(Index).Key Then             ' find key in the keys collection
        CopyMemory ByVal VarPtr(tempKey), m_Images(Index).Key, 4&
        kIndex = local_FindKey(tempKey, True)
        CopyMemory ByVal VarPtr(tempKey), vbDefault, 4&
    End If
    
    If newKey = vbNullString Then
        If kIndex = vbDefault Then Exit Sub ' changing from null to null; no change
    Else
        newKindex = local_FindKey(newKey, bUnique) ' find placement for new key in collection
        If Not bUnique Then
            If Not newKindex = kIndex Then  ' duplicated key
                Err.Raise 1963, "Images.Key", "Duplicated Key. Key already exists in this collection"
                Exit Sub
            Else
                m_Keys(kIndex).Key = newKey ' else changing key case? just do it as requested
                m_Images(Index).Key = StrPtr(m_Keys(kIndex).Key) ' update cross-reference
                Exit Sub
            End If
        End If
    End If
    If kIndex Then
        m_Keys(kIndex).Key = vbNullString   ' release old key
        If kIndex < m_KeyCount Then ' need to shift the old key out of collection
            CopyMemory ByVal VarPtr(m_Keys(kIndex + 1)), ByVal VarPtr(m_Keys(kIndex)), (m_KeyCount - kIndex) * 8&
            CopyMemory ByVal VarPtr(m_Keys(m_KeyCount).Key), vbDefault, 4&
        End If
        m_KeyCount = m_KeyCount - 1
    End If
    If newKindex Then                   ' if replacing, then do that
        m_Images(Index).Key = StrPtr(newKey)
        local_AddKey newKey, vbDefault, ObjPtr(m_Images(Index).ImageClass)
        CopyMemory ByVal VarPtr(newKey), vbDefault, 4&
    Else                                ' else making it null & update cross-reference
        m_Images(Index).Key = vbDefault
        If m_KeyCount + 3 < UBound(m_Keys) Then ReDim Preserve m_Keys(1 To m_KeyCount)
    End If

End Sub

Private Sub local_ReIndexImage(oldIndex As Long, newIndex As Long)

    ' Changes the index of a specific item

    If oldIndex = newIndex Then Exit Sub
    If newIndex > m_ItemCount Or newIndex < 1 Then
        Err.Raise 1962, "Images.Index", "Invalid Key or Index"
        Exit Sub
    End If
    
    Dim tData As ImageDataStructure, oSA As SafeArray
    Dim pIndex As Long, tagPtr As Long
    Dim dibBytes() As Byte, shiftBytes() As Byte
    Dim scanWidth As Long, shiftScanWidth As Long
    Dim X As Long, Y As Long, X2 As Long
    
    If local_AdjustDibSize(vbDefault, vbDefault, vbDefault) = False Then Exit Sub
    ' copy the item that will be replaced during shifting
    CopyMemory ByVal VarPtr(tData), ByVal VarPtr(m_Images(oldIndex)), 8&
    tagPtr = StrPtr(m_Tags(oldIndex))
    
    ' get is index in the pointer-to-index array
    pIndex = local_FindPointer(ObjPtr(m_Images(oldIndex).ImageClass), False, m_ItemCount)
    
    ' bitmap shifting settings
    scanWidth = m_Width * 4&                                ' nr bytes for single image
    shiftScanWidth = Abs(oldIndex - newIndex) * scanWidth   ' nr bytes to shift at once
    ' setup arrays
    iparseOverlayHost_Byte dibBytes(), VarPtr(oSA), 2, m_Height, m_Strip.scanWidth, m_Strip.BitsPointer
    ReDim shiftBytes(0 To scanWidth - 1, 0 To m_Height - 1)
    ' copy the image that will be replaced during shifting
    X2 = (newIndex - 1) * scanWidth
    X = (oldIndex - 1) * scanWidth
    For Y = 0 To m_Height - 1
        CopyMemory shiftBytes(0, Y), dibBytes(X, Y), scanWidth
    Next
    
    If newIndex < oldIndex Then 'shifting to the right
        local_ReIndexPointers newIndex, oldIndex - 1, 1, False
        ' shift the items to the right
        CopyMemory ByVal VarPtr(m_Images(newIndex + 1)), ByVal VarPtr(m_Images(newIndex)), (oldIndex - newIndex) * 8&
        CopyMemory ByVal VarPtr(m_Tags(newIndex + 1)), ByVal VarPtr(m_Tags(newIndex)), (oldIndex - newIndex) * 4&
        ' shift the bitmap to the right
        X = newIndex * scanWidth
        For Y = 0 To m_Height - 1
            CopyMemory dibBytes(X, Y), dibBytes(X - scanWidth, Y), shiftScanWidth
            CopyMemory dibBytes(X2, Y), shiftBytes(0, Y), scanWidth
        Next
            
    Else                        'shifting to the left
        local_ReIndexPointers oldIndex + 1, newIndex, -1, False
        ' shift the items to the left
        CopyMemory ByVal VarPtr(m_Images(oldIndex)), ByVal VarPtr(m_Images(oldIndex + 1)), (newIndex - oldIndex) * 8&
        CopyMemory ByVal VarPtr(m_Tags(oldIndex)), ByVal VarPtr(m_Tags(oldIndex + 1)), (newIndex - oldIndex) * 4&
        ' shift the bitmap to the left
        X = (oldIndex - 1) * scanWidth
        For Y = 0 To m_Height - 1
            CopyMemory dibBytes(X, Y), dibBytes(X + scanWidth, Y), shiftScanWidth
            CopyMemory dibBytes(X2, Y), shiftBytes(0, Y), scanWidth
        Next
    End If
    iparseOverlayHost_Byte dibBytes(), 0, 0, 0, 0, 0
    
    ' put newIndex in its new spot & update index
    CopyMemory ByVal VarPtr(m_Images(newIndex)), ByVal VarPtr(tData), 8&
    CopyMemory ByVal VarPtr(m_Tags(newIndex)), tagPtr, 4&
    FillMemory ByVal VarPtr(tData), 8&, 0
    m_Indexes(pIndex).Index = newIndex
    

End Sub

Private Sub local_AddPointer(oPtr As Long, oIndex As Long, bAppending As Boolean)

    ' Adds a new item to the pointer-to-index array

    Dim pIndex As Long
    pIndex = local_FindPointer(oPtr, True, m_ItemCount + bAppending)
    
    If pIndex < m_ItemCount Then
        CopyMemory ByVal VarPtr(m_Indexes(pIndex + 1)), ByVal VarPtr(m_Indexes(pIndex)), (m_ItemCount - pIndex) * 8&
    End If
    m_Indexes(pIndex).ImagePtr = oPtr
    m_Indexes(pIndex).Index = oIndex
    
End Sub


Private Function local_AdjustDibSize(newCount As Long, ByVal adjIndex As Long, ShiftSlots As Long) As Boolean
        
    ' Helper function. Manages size of imagelist in response to additions/deletions
    ' Parameters:
    '   newCount :: the number of image slots required
    '   adjIndex :: negative is index being deleted, positive is index of insertion
    '               zero means to trim, being exported
        
    Dim bResize As Boolean
    Dim bInIDE As Boolean
    Dim aSrc() As Byte, aDst() As Byte
    Dim oSA As SafeArray, oDA As SafeArray
    Dim Y As Long
    Dim fromSlot As Long, toSlot As Long, nrSlots As Long
    Dim scanWidth As Long, shiftScanWidth As Long
    Dim tDIB As c32bppDIB
    
    ' check for delay load. m_Strip will be created, but not initialized
    ' When ImageType=imgError, this should be the case
    If m_Strip.ImageType = imgError Then
        ' see if image waiting to be processed
        bResize = m_Strip.LoadPicture_FromOrignalFormat()
        If bResize = False Then
            ' see if our m_Images array is empty, shoudln't be if DelayLoad=True
            If iparseIsArrayEmpty(VarPtrArray(m_Images)) Then
                ' create the blank image
                bResize = m_Strip.InitializeDIB(m_Width * UBound(m_Images), m_Height)
            End If
        Else
            m_Strip.SetOriginalFormat aSrc()    ' don't need cached bytes any long
        End If
        If bResize = False Then 'failure, abort
            Me.Clear
            Err.Raise 1977, "Images", "Low system resource preventing appending additional images"
            Exit Function
        End If
    End If
    If newCount = vbDefault Then ' can be 0 only if a routine wants to ensure loaded if delayload=true
        local_AdjustDibSize = True
        Exit Function
    End If
    nrSlots = m_Strip.Width \ m_Width
    If adjIndex = vbDefault Then
        If (nrSlots = newCount) Then
            local_AdjustDibSize = True
            Exit Function
        End If
        bResize = True
    Else
        bResize = (nrSlots < newCount)
    End If
    
    If bResize Then ' need different sized dib
        Set tDIB = New c32bppDIB
        tDIB.InitializeDIB newCount * m_Width, m_Height
        nrSlots = newCount
        iparseOverlayHost_Byte aDst(), VarPtr(oDA), 2, m_Height, tDIB.scanWidth, tDIB.BitsPointer
    Else    ' source & destination are the same
        iparseOverlayHost_Byte aDst(), VarPtr(oDA), 2, m_Height, m_Strip.scanWidth, m_Strip.BitsPointer
    End If
    ' overlay source
    iparseOverlayHost_Byte aSrc(), VarPtr(oSA), 2, m_Height, m_Strip.scanWidth, m_Strip.BitsPointer
    
    scanWidth = m_Width * 4&
    If adjIndex = vbDefault Then ' trimming in prepartion for export
        shiftScanWidth = newCount * scanWidth
        For Y = 0 To m_Height - 1
            CopyMemory aDst(vbDefault, Y), aSrc(vbDefault, Y), shiftScanWidth
        Next
    Else
        If ShiftSlots > vbDefault Then ' inserting from slot adjIndex
            ' when inserting we need need to transfer images up to adjIndex only if a new DIB was created
            If Not tDIB Is Nothing Then
                shiftScanWidth = (adjIndex - 1) * scanWidth
                If shiftScanWidth > vbDefault Then
                    For Y = 0 To m_Height - 1
                        CopyMemory aDst(vbDefault, Y), aSrc(vbDefault, Y), shiftScanWidth
                    Next
                End If
            End If
            ' and now to calc byte offsets to copy from & copy to
            fromSlot = (adjIndex - 1) * scanWidth
            toSlot = (adjIndex + ShiftSlots - 1) * scanWidth
            shiftScanWidth = (nrSlots - adjIndex - ShiftSlots + 1) * scanWidth
        Else                         ' removing from slot adjIndex
            ' when removing, no new DIB is created, so we don't need to calc xferScanWidth
            ' but calc byte offsets to copy from and copy to
            toSlot = (adjIndex + ShiftSlots) * scanWidth
            fromSlot = adjIndex * scanWidth
            shiftScanWidth = (nrSlots - adjIndex) * scanWidth
        End If
        ' last is to calc nr of bytes to shift right or left
        
        If shiftScanWidth > vbDefault Then
            For Y = 0 To m_Height - 1
                CopyMemory aDst(toSlot, Y), aSrc(fromSlot, Y), shiftScanWidth
            Next
        End If
        If ShiftSlots > vbDefault Then              ' inserting, erase the insert area for a clean slate
            shiftScanWidth = scanWidth * ShiftSlots
            For Y = 0 To m_Height - 1
                FillMemory aDst(fromSlot, Y), shiftScanWidth, 0
            Next
        End If
    End If
    iparseOverlayHost_Byte aSrc(), 0, 0, 0, 0, 0
    iparseOverlayHost_Byte aDst(), 0, 0, 0, 0, 0
    
    If Not tDIB Is Nothing Then m_Strip.SwapDIBs tDIB
    
    local_AdjustDibSize = True
        
End Function

Private Function local_PointerToIndex(oPtr As Long) As Long
    ' function to convert item pointer to item index
    Dim pIndex As Long
    pIndex = local_FindPointer(oPtr, True, m_ItemCount)
    local_PointerToIndex = m_Indexes(pIndex).Index
End Function

Private Function local_PointerToKey(oPtr As Long) As String
    ' function to convert item pointer to its key
    Dim pKey As String
    Dim pIndex As Long
    pIndex = local_PointerToIndex(oPtr)
    If Not m_Images(pIndex).Key = vbDefault Then
        CopyMemory ByVal VarPtr(pKey), m_Images(pIndex).Key, 4&
        local_PointerToKey = pKey
        CopyMemory ByVal VarPtr(pKey), vbDefault, 4&
    End If
End Function

Private Function local_KeyToObject(Key As Variant, Optional itemIndex As Long) As cListImage

    ' converts a variant index or key to cListImage class

    Dim kIndex As Long, bIsNew As Boolean
    
    Select Case VarType(Key)
    Case vbString                           ' Strings are always treated as Keys
        If Not Key = vbNullString Then
            kIndex = local_FindKey(CStr(Key), bIsNew)
            If bIsNew = False Then
                itemIndex = local_PointerToIndex(m_Keys(kIndex).ImagePtr)
                Set local_KeyToObject = m_Images(itemIndex).ImageClass
            End If
        End If
    Case vbLong, vbInteger, vbDouble, vbByte, vbSingle ' Numerical varTypes are treated as Indexes
        On Error Resume Next
        kIndex = Int(Key)
        If Err Then
            Err.Clear
        Else
            If kIndex > vbDefault And kIndex <= m_ItemCount Then
                itemIndex = kIndex
                Set local_KeyToObject = m_Images(kIndex).ImageClass
            End If
        End If
    End Select

End Function

Private Function local_FindKey(Key As String, ByRef isNew As Boolean) As Long

    ' BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between just 1 and 20 iterations
    
    ' [in] Key. A value to search for. Order is always ascending
    ' [out] isNew. If Key not found, isNew is True else False
    ' [out] Return value: The Index where Key was found or where the new Key should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If m_KeyCount = 0& Then
        newIndex = 1&
        isNew = True
    Else
    
        UB = m_KeyCount
        LB = 1&
        
        Do Until LB > UB
            newIndex = LB + ((UB - LB) \ 2&)
            Select Case StrComp(Key, m_Keys(newIndex).Key, vbTextCompare)
            Case 0 ' key found
                Exit Do
            Case Is < 0 ' key is < current key
                UB = newIndex - 1&
            Case Else ' key is > current key
                LB = newIndex + 1&
            End Select
        Loop
    
        If LB > UB Then  ' key was not found
            If StrComp(Key, m_Keys(newIndex).Key, vbTextCompare) > 0 Then newIndex = newIndex + 1
            isNew = True
        Else
            isNew = False
        End If
    
    End If
    
    local_FindKey = newIndex

End Function

Private Function local_FindPointer(oPtr As Long, isNew As Boolean, imgCount As Long) As Long

    ' Binary search looking in the pointer structure. See local_FindKeys for algorithm logic remarks
    
    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If imgCount = 0& Then
        newIndex = 1&
        isNew = True
    Else
    
        LB = 1&
        UB = imgCount
        
        Do Until LB > UB
            newIndex = LB + ((UB - LB) \ 2&)
            Select Case oPtr - m_Indexes(newIndex).ImagePtr
            Case 0 ' key found
                Exit Do
            Case Is < 0 ' key is < current key
                UB = newIndex - 1&
            Case Else ' key is > current key
                LB = newIndex + 1&
            End Select
        Loop
    
        If LB > UB Then  ' key was not found
            If oPtr > m_Indexes(newIndex).ImagePtr Then newIndex = newIndex + 1
            isNew = True
        Else
            isNew = False
        End If
    
    End If
    
    local_FindPointer = newIndex

End Function

Private Function local_IsUnCompiled(bInIDE As Boolean) As Boolean
    bInIDE = True
    local_IsUnCompiled = True
End Function


Private Function local_ExtractIcon(srcList As c32bppDIB, ByVal iIndex As Long, pBitDepth As Long, ByVal imgWidth As Long, ByVal imgHeight As Long) As Long

    Dim bUnique As Boolean, bPerPixelAlpha As Boolean
    Dim X As Long, Y As Long
    Dim bIcon() As Byte, Pow2(0 To 8) As Integer
    Dim bSource() As Byte, oSA As SafeArray
    Dim lSource() As Long, oSALong As SafeArray                             ' source overlay use (arrays never initialized)
    Dim cPalette(1 To 256) As Long, ColorCount As Long, palColor As Byte    ' palettized icon use
    Dim scanWidth As Long, pScanWidth As Long, mScanWidth As Long           ' source, palette, mask scan widths
    Dim pIndex As Long, aPtr As Long                                        ' array pointers
    Dim pOffset As Long, mOffset As Long                                    ' palette & mask array offsets
    Dim maskByte As Byte, maskPos As Long                                   ' bit shifting use
    Dim resLength As Long                                                   ' icon bit depth & overall size
    
    Pow2(0) = 1 ' create a power of two lookup table
    For X = 1 To 8: Pow2(X) = Pow2(X - 1) * 2: Next
    
    scanWidth = imgWidth * 4&                            ' scanwidth of source image
    mScanWidth = iparseByteAlignOnWord(1, imgWidth)      ' scanwidth for dest mask bytes
    iIndex = iIndex - 1
    
    ' presize the array to max expected size; will be significantly smaller if image can be paletted
    mOffset = scanWidth * imgHeight + 39 + 1024          ' identify where in array the maskbytes will begin
    ReDim bIcon(0 To mOffset + mScanWidth * imgHeight)   ' oversize array vs re-dimming later
    
    ' create a byte overlay array so we can look at the alpha bytes only
    iparseOverlayHost_Byte bSource(), VarPtr(oSA), 2, srcList.Height, srcList.scanWidth, srcList.BitsPointer
    ' add Long array overlay so we can easily look at 4 byte colors
    iparseOverlayHost_Long lSource(), VarPtr(oSALong), 2, srcList.Height, srcList.Width, srcList.BitsPointer
    
    
    ' create mask and determine if paletting is possible
    aPtr = iIndex * scanWidth + 3                       ' position for first alpha byte in source
    For Y = 0 To imgHeight - 1
        maskPos = 7                                     ' position for 1st mask bit
        pIndex = mOffset + (mScanWidth * Y)             ' position in array for this row of source pixels
        For X = aPtr To aPtr + scanWidth - 4 Step 4&
            If bSource(X, Y) = 0 Then                   ' set bit to 1 if pixel is transparent
                maskByte = maskByte Or Pow2(maskPos)
            Else                                        ' determine if per-pixel alphablending is in use
                If bSource(X, Y) < 255 Then bPerPixelAlpha = True   ' will be converted to 32bpp icon
            End If
            If maskPos = 0 Then                         ' mask shifted out
                bIcon(pIndex) = maskByte                ' append mask byte
                pIndex = pIndex + 1                     ' move to next mask index
                maskPos = 7                             ' reset mask bit pos
                maskByte = 0                            ' reset byte
            Else
                maskPos = maskPos - 1                   ' shift bit mask
            End If
        Next
        If maskByte Then                                ' widths not evenly divisible by 4 will have unfilled bits
            bIcon(pIndex) = maskByte                    ' write the final byte for this row
            maskByte = 0
        End If
    Next
    
    ' user requests paletted icon, see if it is doable; not doable if bSimpleAlpha=False
    aPtr = iIndex * imgWidth                            ' position of of first source pixel
    If Not bPerPixelAlpha Then
        ' here we are going to count colors, if we exceed 256 then we abort and cannot palettize
        For Y = 0 To imgHeight - 1
            For X = aPtr To aPtr + imgWidth - 1         ' log each color
                pIndex = local_FindColor(cPalette, lSource(X, Y), ColorCount, bUnique)
                If bUnique Then                         ' new color for the palette
                    If ColorCount = 256 Then            ' exceeded 256
                        ColorCount = 0&                 ' flag indicating can't palettize
                        Y = imgHeight                   ' aborts outer loop
                        Exit For
                    Else
                        ColorCount = ColorCount + 1     ' increment palette count
                        If pIndex < ColorCount Then     ' keep palette sorted for searching
                            CopyMemory cPalette(pIndex + 1), cPalette(pIndex), (ColorCount - pIndex) * 4&
                        End If
                        cPalette(pIndex) = lSource(X, Y) ' add color to palette
                    End If
                End If
            Next
        Next
    End If
    
    ' now build the color portion of the icon
    If ColorCount = 0& Then
    
        pIndex = 40                                     ' position in icon array where color data starts
        If bPerPixelAlpha Then                          ' can create true color icon (supported by Win2K+)

' 32bpp icons
            pBitDepth = 32                              ' icon bit depth & size of color data needed
            pScanWidth = scanWidth
            aPtr = iIndex * scanWidth                   ' position of of first source pixel (Byte)
            ' this should be the easiest, but we need to remove pre-multiplication
            For Y = 0& To imgHeight - 1&
                For X = aPtr To aPtr + scanWidth - 1 Step 4
                    Select Case bSource(X + 3&, Y)
                    Case 255
                        CopyMemory bIcon(pIndex), bSource(X, Y), 4&
                    Case 0 ' do nothing
                    Case Else
                        palColor = bSource(X + 3&, Y)
                        bIcon(pIndex) = (255& * bSource(X, Y) \ palColor)
                        bIcon(pIndex + 1&) = (255& * bSource(X + 1&, Y) \ palColor)
                        bIcon(pIndex + 2&) = (255& * bSource(X + 2&, Y) \ palColor)
                        bIcon(pIndex + 3&) = palColor
                    End Select
                    pIndex = pIndex + 4&
                Next
            Next
            aPtr = iIndex * imgWidth                    ' position of of first source pixel (Long)
        Else

' 24bpp Icons
            pBitDepth = 24                              ' icon bit depth & calc size of color data needed
            pScanWidth = iparseByteAlignOnWord(pBitDepth, imgWidth)
            For Y = 0 To imgHeight - 1
                pIndex = 40 + Y * pScanWidth            ' move to position for current row
                For X = aPtr To aPtr + imgWidth - 1
                    CopyMemory bIcon(pIndex), lSource(X, Y), 3& ' copy 3 color bytes & increment pointer
                    pIndex = pIndex + 3
                Next
            Next
        End If
        resLength = pScanWidth * imgHeight + mScanWidth * imgHeight + 40  ' calc actual size of icon data to be used
        ' shift the mask bytes to their proper array position
        CopyMemory bIcon(pScanWidth * imgHeight + 40), bIcon(mOffset), mScanWidth * imgHeight
        
    Else    ' we can palettize, create the palette indexes
        
        If ColorCount > 16 Then
' 8bpp icons
             pBitDepth = 8              ' use 256 color palette (2^8)
        ElseIf ColorCount > 2 Then
' 4bpp icons
            pBitDepth = 4               ' use 16 color palette (2^4)
        Else
' 1bpp icons
            pBitDepth = 1               ' use 2 color palette  (2^1)
        End If
' 2bpp icons can be created, but Windows doesn't seem to support them. pBitDepth=2
        
        pScanWidth = iparseByteAlignOnWord(pBitDepth, imgWidth)      ' palette indexes scanwidth
        pOffset = Pow2(pBitDepth) * 4 + 40                           ' start of pixel data (after palette & bitmapinfo header)
        pIndex = pScanWidth * imgHeight + pOffset                    ' start of mask data
        CopyMemory bIcon(pIndex), bIcon(mOffset), mScanWidth * imgHeight ' shift mask bytes to their final position
        resLength = pIndex + mScanWidth * imgHeight                  ' calc actual size of icon data to be used
        
        ' build the palette indexes
        For Y = 0 To imgHeight - 1
            pIndex = pOffset + (pScanWidth * Y)                     ' position for next row of palette indexes
            ' handle 8bpp separate, less code needed
            If pBitDepth = 8 Then
                For X = aPtr To aPtr + imgWidth - 1
                    bIcon(pIndex) = local_FindColor(cPalette, lSource(X, Y), ColorCount, False) - 1
                    pIndex = pIndex + 1
                Next
            Else            ' 1 or 4 bit icons
                maskPos = 8 - pBitDepth                             ' position in palette's bit mask
                For X = aPtr To aPtr + imgWidth - 1
                    ' add to the palette index byte
                    palColor = local_FindColor(cPalette, lSource(X, Y), ColorCount, False) - 1
                    maskByte = maskByte Or (Pow2(maskPos) * palColor)
                    If maskPos = 0 Then                             ' palette shifted out
                        bIcon(pIndex) = maskByte                    ' append palette index
                        maskPos = 8 - pBitDepth                     ' reset palette bit pos
                        maskByte = 0                                ' reset palette byte
                        pIndex = pIndex + 1                         ' move to next palette index
                    Else
                        maskPos = maskPos - pBitDepth               ' shift bit mask
                    End If
                Next
                If maskByte Then                                    ' write remainder if any
                    bIcon(pIndex) = maskByte
                    maskByte = 0
                End If
            End If
        Next
        
        CopyMemory bIcon(40), cPalette(1), Pow2(pBitDepth) * 4      ' transfer palette to icon array
        ' below is the bitmapinfo header's biClrUsed & biClrImportant members
        ' but browsing icon formats, this does not seem to be needed
        CopyMemory bIcon(32), Pow2(pBitDepth), 2&           ' colors used
        CopyMemory bIcon(36), ColorCount, 4&                ' colors important
    
    End If
    iparseOverlayHost_Byte bSource(), 0, 0, 0, 0, 0         ' remove overlays
    iparseOverlayHost_Long lSource(), 0, 0, 0, 0, 0
    
    ' build the bitmapinfo header
    CopyMemory bIcon(0), 40&, 4&                            ' header size
    CopyMemory bIcon(4), imgWidth, 4&                       ' width
    CopyMemory bIcon(8), CLng(imgHeight * 2), 4&            ' height * 2
    CopyMemory bIcon(12), 1, 2&                             ' planes
    CopyMemory bIcon(14), pBitDepth, 2&                     ' bitcount
    CopyMemory bIcon(20), CLng(resLength - 40), 4&          ' color bits + mask bits + palette
        
    ' return result
    local_ExtractIcon = CreateIconFromResourceEx(bIcon(0), resLength, 1, &H30000, 0, 0, 0&)


End Function

Private Function local_FindColor(ByRef PaletteItems() As Long, Color As Long, Count As Long, isNew As Boolean) As Long

    ' MODIFIED BINARY SEARCH ALGORITHM -- Divide and conquer.

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If Count = 0& Then
        local_FindColor = 1&
        isNew = True
        Exit Function
    End If
    
    UB = Count
    LB = 1&
    
    Do Until LB > UB
        newIndex = LB + ((UB - LB) \ 2&)
        If PaletteItems(newIndex) = Color Then
            Exit Do
        ElseIf PaletteItems(newIndex) > Color Then ' new color is lower in sort order
            UB = newIndex - 1&
        Else ' new color is higher in sort order
            LB = newIndex + 1&
        End If
    Loop

    If LB > UB Then  ' color was not found
        If Color > PaletteItems(newIndex) Then newIndex = newIndex + 1&
        isNew = True
    Else
        isNew = False
    End If
    
    local_FindColor = newIndex

End Function


Private Sub Class_Initialize()
    Set m_Strip = New c32bppDIB
End Sub

Private Sub Class_Terminate()
    Me.Clear
    Set m_Strip = Nothing
    Set m_Slot = Nothing
End Sub

' Call back routines from the individual child image classes

Private Property Let IItemData_ClassDataLong(DataName As eDataNames, oID As Long, ByVal newValue As Long)
    ' The only non-string setting that child class can set is the Index
    If DataName = eData_Index Then
        local_ReIndexImage local_PointerToIndex(oID), newValue
    ElseIf DataName = eData_EnumReset Then
        m_EnumIndex = vbDefault
    End If
End Property
Private Property Get IItemData_ClassDataLong(DataName As eDataNames, oID As Long) As Long
    ' return width, height,or index for the calling item
    Select Case DataName
        Case eData_Height: IItemData_ClassDataLong = m_Height
        Case eData_Width: IItemData_ClassDataLong = m_Width
        Case eData_Index: IItemData_ClassDataLong = local_PointerToIndex(oID)
        Case eData_EnumNext
            m_EnumIndex = m_EnumIndex + 1
            IItemData_ClassDataLong = (m_EnumIndex <= m_ItemCount)
    End Select
End Property

Private Property Get IItemData_ClassDataObject(DataName As eDataNames, oID As Long) As Variant
    ' Note: Kelly Ethridge example uses VariantCopy API to transfer enumerated item
    ' to property; however, no can do with objects it appears -- crash. This works.
    If DataName = eData_EnumItem Then Set IItemData_ClassDataObject = m_Images(m_EnumIndex).ImageClass
End Property

Private Property Let IItemData_ClassDataString(DataName As eDataNames, oID As Long, ByVal newValue As String)
    Dim iIndex As Long
    If DataName = eData_Key Then            ' change item's key
        iIndex = local_PointerToIndex(oID)
        If iIndex Then local_ReyKeyImage iIndex, newValue
    ElseIf DataName = eData_Tag Then        'change item's tag
        iIndex = local_PointerToIndex(oID)
        If iIndex Then m_Tags(iIndex) = newValue
    End If
End Property
Private Property Get IItemData_ClassDataString(DataName As eDataNames, oID As Long) As String
    ' retrieve items' key or tag
    Dim iIndex As Long
    If DataName = eData_Key Then
        IItemData_ClassDataString = local_PointerToKey(oID)
    ElseIf DataName = eData_Tag Then
        iIndex = local_PointerToIndex(oID)
        If iIndex Then IItemData_ClassDataString = m_Tags(iIndex)
    End If
End Property


' class to class use-only functions

Friend Sub xcom_InitializeImages(pParent As IItemData, Width As Long, Height As Long, initCount As Long)
    Set mParent = pParent
    If initCount Then
        ReDim m_Images(1 To initCount)
        ReDim m_Indexes(1 To initCount)
        ReDim m_Tags(1 To initCount)
    End If
    m_Width = Width
    m_Height = Height
End Sub


' Imports imagelist data from a byte array from usercontrol's ReadProperty
' Also called by usercontrol's ImportImageList function
Friend Function xcom_ImportList(bImport() As Byte, HeaderBytes As Long, ValidationOffset As Long, iVersion As Long) As Boolean

    Dim lPointer As Long, lBoundOffset As Long
    Dim iItem As Long
    Dim Key As String, strLen As Long, keyLen As Long
    Dim bInIDE As Boolean
    Dim bCompressMethod As Byte ' 0=either PNG or raw Bimtap, 1=compressed bits only, 2=compressed bmp w/header, 4=compressed keys/tags
    Dim xtraBytes As Long
    Dim bOverlay() As Byte, oSA As SafeArray
    Dim bKeys() As Byte
    
    Dim cCompressor As cHuffman
    Dim tImage As cListImage
    
    ' Import data is same as described in ExportImageList
    ' Nr Bytes  Data
    '   nnn     header bytes used by the usercontrol (image list name, version, other as needed)
    '   4       nrItems :: overall DIB width/nrItems = individual image width.  ImageHeight is overall DIB height
    '   4       nrKeys  :: speeds up importing because we can pre-initialize key arrays
    '   8       width & height of images
    '   for each item, following is defined
    '       4   length of key & tag (zero if not keyed/tagged) HiWord=Tag length, LoWord=Key Length
    '       nnn key in unicode if present
    '       nnn tag in unicode of present
    '   4       extra offset bytes to beginning of image. 2 for Version 1
    '   nnn     extra bytes if used
    '   ---     image begins
    
    Me.Clear
    
    lPointer = LBound(bImport) + HeaderBytes                ' start of array
    lBoundOffset = LBound(bImport) - 1
    m_KeyCount = vbDefault
    CopyMemory m_ItemCount, bImport(lPointer), 4&         ' number of sub images?
    If m_ItemCount Then
        ReDim m_Images(1 To m_ItemCount)
        ReDim m_Indexes(1 To m_ItemCount)
        ReDim m_Tags(1 To m_ItemCount)
        CopyMemory iItem, bImport(lPointer + 4), 4&         ' number of keyed images?
        If iItem Then
            ReDim m_Keys(1 To iItem)                        ' resize arrays
        End If
        CopyMemory m_Width, bImport(lPointer + 8), 4&       ' width of images
        CopyMemory m_Height, bImport(lPointer + 12), 4&     ' height of images
        
        ' get version data...
        CopyMemory xtraBytes, bImport(ValidationOffset + lBoundOffset), 4&
        If iVersion > 1 Then    ' we have image/key/tag compression options
            CopyMemory bCompressMethod, bImport(ValidationOffset + 6 + lBoundOffset), 1&
            If (bCompressMethod And 4) = 4 Then ' keys/tags are huffmann compressed
                Set cCompressor = New cHuffman
                ' compressed bytes will always start 16bytes from the header bytes
                cCompressor.DecodeByte bImport, lPointer + 16, bKeys(), Nothing
                ' overlay array on returned array; neat hack in using safearray as "Any" array
                iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, Abs(UBound(bKeys) - LBound(bKeys)) + 1, 0, VarPtr(bKeys(LBound(bKeys)))
                lPointer = 0    ' overlays are always zero-bound
            End If
        End If
        If cCompressor Is Nothing Then  ' keys/tags are not compressed
            ' overlay array over our import data
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, Abs(UBound(bImport) - LBound(bImport)) + 1, 0, VarPtr(bImport(LBound(bImport)))
            lPointer = HeaderBytes + 16 ' where keys/tags begin
        End If
        
        ' routine extracts data from 1 of 2 arrays, the overlay array points to either
        For m_ItemCount = 1 To m_ItemCount
            Set tImage = New cListImage
            CopyMemory strLen, bOverlay(lPointer), 4&        ' key, tag?
            keyLen = (strLen And &HFFFF&)                   ' key length (max is 512 or 256 characters)
            If keyLen Then
                Key = Space$(keyLen \ 2)                    ' build string to hold key
                CopyMemory ByVal StrPtr(Key), bOverlay(lPointer + 4), keyLen  ' copy string into the key
                If m_KeyCount = 0& Then                     ' handle 1st instance, else local_AddKey will reset array
                    m_KeyCount = 1&                         ' update key count
                    m_Keys(m_KeyCount).Key = Key            ' copy the key
                    m_Keys(m_KeyCount).ImagePtr = ObjPtr(tImage)
                    m_Images(m_ItemCount).Key = StrPtr(m_Keys(m_KeyCount).Key) ' cross-reference
                Else
                    local_AddKey Key, vbDefault, ObjPtr(tImage) ' call function to sort key & add references
                    m_Images(m_ItemCount).Key = StrPtr(Key)
                    CopyMemory ByVal VarPtr(Key), vbDefault, 4&
                End If
            End If
            strLen = (strLen \ &H10000)                     ' tag length is in HiWord of value
            If strLen Then                                  ' has a tag property?
                m_Tags(m_ItemCount) = Space$(strLen \ 2)    ' create string for the tag
                CopyMemory ByVal StrPtr(m_Tags(m_ItemCount)), bOverlay(lPointer + keyLen + 4), strLen ' copy string into the tag
            End If
            local_AddPointer ObjPtr(tImage), m_ItemCount, True
            tImage.xcom_InitializeItem Me
            Set m_Images(m_ItemCount).ImageClass = tImage
            lPointer = lPointer + keyLen + strLen + 4
        Next
        m_ItemCount = m_ItemCount - 1
        ' done, remove the overlay
        iparseOverlayHost_Byte bOverlay(), 0, 0, 0, 0, 0
        
        If Not cCompressor Is Nothing Then
            ' keys/tags were uncompressed; free array & update current compression options
            bCompressMethod = (bCompressMethod And Not 4)
            Erase bKeys
        End If
        lPointer = ValidationOffset + lBoundOffset ' pointer to xtraBytes data
        ' Version 0/1: 2 bytes
            CopyMemory m_DelayLoad, bImport(lPointer + 4), 2&
        ' Version 2: 1 byte
            ' processed above; contains image/key/tag compression if used
        ' Version 3+: process those bytes
        '  ...
        lPointer = lPointer + xtraBytes + 4     ' move to beginning of image data
        
        ' one subversion of version 2 compresses only pixel data. The other compresses entire bitmap with header
        ' When entire bitmap is compressed, we can allow DelayLoad to be active else we can't
        If bCompressMethod = 2 Then
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            If cCompressor.DecodeByte(bImport(), lPointer, bOverlay(), Nothing) = False Then
                Me.Clear
                Exit Function
            End If
            Set cCompressor = Nothing
        End If
        
        ' create the DIB class and import the image, unless delay load is set
        Set m_Strip = New c32bppDIB
        m_Strip.gdiToken = mParent.ClassDataLong(eData_GDItoken, vbDefault)
        Debug.Assert local_IsUnCompiled(bInIDE)
        m_Strip.KeepGDIplusActive = Not bInIDE
        
        If m_DelayLoad = True And bCompressMethod <> 1 Then
            ' delay load does not create the 32bpp imagelist until it is first used
            ' Resize the array so it only holds the image data & format header
            If bCompressMethod = vbDefault Then
                ReDim Preserve bOverlay(1 To UBound(bImport) - lPointer + 1)
                CopyMemory bOverlay(1), bImport(lPointer), UBound(bOverlay)
            End If
            ' have class maintian the array until needed
            m_Strip.SetOriginalFormat bOverlay
            xcom_ImportList = True
            
        ElseIf bCompressMethod = 1 Then
            ' compressed image bytes only. When uncompressed, done directly to DIB
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            xcom_ImportList = cCompressor.DecodeByte(bImport(), lPointer, bImport(), m_Strip)
            Set cCompressor = Nothing
            
        Else
            ' no delay load, create DIB now
            If bCompressMethod = 2 Then
                xcom_ImportList = m_Strip.LoadPicture_Stream(bOverlay(), , , 1, UBound(bOverlay))
            Else
                xcom_ImportList = m_Strip.LoadPicture_Stream(bImport(), , , lPointer, UBound(bImport) - lPointer + 1)
            End If
        End If
        
        ' when compiled, add this efficiency enhancement. Can crash if turned on in IDE
        mParent.ClassDataLong(eData_Height, vbDefault) = m_Height
        mParent.ClassDataLong(eData_Width, vbDefault) = m_Width
    End If
    
End Function


' Exports imagelist data to a byte array for storage from usercontrol's WriteProperty
' Also called by usercontrol's ExportImageList function
Friend Function xcom_ExportList(bExport() As Byte, HeaderBytes As Long, _
                            ValidationOffset As Long, _
                            ImageOffset As Long, _
                            ByVal CompressOption As Boolean) As Boolean

    Dim lSizeArray As Long, lImageSize As Long
    Dim lPointer As Long
    Dim iItem As Long
    Dim bImage() As Byte, bKeys() As Byte
    Dim oSA As SafeArray, bOverlay() As Byte
    Dim Key As String
    Dim keyLen As Long, tagLen As Long
    Dim bCompressMethod As Byte ' 0=PNG, or uncompressed BMP , 1=BMP compressed
    Dim cCompressor As cHuffman
    
    Const xtraBytes As Long = 3
    
    ' ^^  version 1: delay load only extra property
    ' ^^  version 2: + 1 byte: 0=png format, 1=huffmann compressed bitmap, 2=uncompressed bitmap
    
    ' Export data will look like following:
    ' Nr Bytes  Data
    '   nnn     header bytes used by the usercontrol (image list name, version, other as needed)
    '   4       nrItems :: overall DIB width/nrItems = individual image width.  ImageHeight is overall DIB height
    '   4       nrKeys  :: speeds up importing because we can pre-initialize key arrays
    '   8       width & height of images
    '   for each item, following is defined
    '       4   length of key & tag (zero if not keyed/tagged) HiWord=Tag length, LoWord=Key Length
    '       nnn key in unicode if present
    '       nnn tag in unicode of present
    '   4       extra offset bytes to beginning of image. Validation Offset
    '   nnn     extra bytes if used
    '   ---     image begins. Image Offset
    
    If m_ItemCount Then
        
        ' save imagelist to PNG/BMP in the bImage() array
        If m_Strip.GetOrginalFormat(bImage) = False Then
            Call local_AdjustDibSize(m_ItemCount, vbDefault, vbDefault) ' trim the imagelist to remove extra space if needed
            If m_Strip.SaveToStream_PNG(bImage) = False Then    ' if no GDI+ or zLib installed, then
                If m_Strip.SaveToStream(bImage) = False Then Exit Function
                ' to compress or not?
                If CompressOption Then   ' yep try it
                    Set cCompressor = New cHuffman
                    If cCompressor.EncodeListData(Nothing, bImage, bOverlay) Then
                        Erase bImage
                        CopyMemory ByVal VarPtrArray(bImage), ByVal VarPtrArray(bOverlay), 4&
                        CopyMemory ByVal VarPtrArray(bOverlay), vbDefault, 4&
                        bCompressMethod = 2
                    End If
                End If
            End If
        End If
        lImageSize = UBound(bImage) - LBound(bImage) + 1
        
        ' collecting key/tag info. Loop thru once to count total size of all
        For iItem = 1 To m_ItemCount  ' start counting how many bytes are needed for non-image data
            If m_Images(iItem).Key Then
                CopyMemory ByVal VarPtr(Key), m_Images(iItem).Key, 4& ' create soft reference to key -- must be removed by end of routine
                lSizeArray = lSizeArray + Len(Key) * 2  ' unicode size of key
            End If
            ' add space for the tag if it is used
            If m_Tags(iItem) <> vbNullString Then
                lSizeArray = lSizeArray + Len(m_Tags(iItem)) * 2 ' unicode size of key
            End If
        Next
        
        If lSizeArray > 1024 Then       ' compressing less than 1kb generally not much savings
            bCompressMethod = (bCompressMethod Or 4)        ' add flag to show compressed
            lSizeArray = lSizeArray + m_ItemCount * 4      ' calc size of array needed
            ReDim bKeys(1 To lSizeArray)                    ' size the array
            ' overlay array over bKeys. This allows the writing to occur to this array
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, lSizeArray, 0, VarPtr(bKeys(1))
            lPointer = 0 ' bOverlay is zero-bound
        End If
        If (bCompressMethod And 4) = 0 Then ' no compression of key/tag data
            ' calc size of array needed for the entire export
            lSizeArray = lSizeArray + HeaderBytes + m_ItemCount * 4 + xtraBytes + 20
            ' ^^ (20= 4: item count, 4:key count, 4:width, 4:height, 4:xtraByte entry)
            ReDim bExport(1 To lSizeArray + lImageSize) ' now we can resize our array
            ' overlay array over bExport. Allows writing to occur to this array
            iparseOverlayHost_Byte bOverlay(), VarPtr(oSA), 1, lSizeArray + lImageSize, 0, VarPtr(bExport(1))
            lPointer = HeaderBytes + 16 ' bOverlay is zero-bound
        End If
        
        ' loop thru key/tags again, this time writing them to bOverlay
        For iItem = 1 To m_ItemCount
            If m_Images(iItem).Key Then                                      ' keyed?
                CopyMemory ByVal VarPtr(Key), m_Images(iItem).Key, 4&        ' create soft reference to key -- must be removed by end of routine
                keyLen = Len(Key) * 2                                       ' cache length & then store key
                CopyMemory bOverlay(lPointer + 4), ByVal m_Images(iItem).Key, keyLen
            Else ' no key
                keyLen = 0&
            End If
            If m_Tags(iItem) <> vbNullString Then                     ' tagged?
                tagLen = Len(m_Tags(iItem)) * 2                        ' cache length & then store tag
                CopyMemory bOverlay(lPointer + keyLen + 4), ByVal StrPtr(m_Tags(iItem)), tagLen
            Else
                tagLen = 0&
            End If
            CopyMemory bOverlay(lPointer), (keyLen Or (&H10000 * tagLen)), 4& ' store lengths of key/tag
            lPointer = lPointer + tagLen + keyLen + 4                        ' position to next available byte
        Next
        iparseOverlayHost_Byte bOverlay(), 0, 0, 0, 0, 0    ' remove overlay
        CopyMemory ByVal VarPtr(Key), 0&, 4&                ' remove soft reference inside the Key variable
        
        If (bCompressMethod And 4) = 4 Then ' we are going to compress the keys/tags
            If cCompressor Is Nothing Then Set cCompressor = New cHuffman
            cCompressor.EncodeListData Nothing, bKeys(), bOverlay()
            ' calculate entire size of the export array
            lSizeArray = HeaderBytes + UBound(bOverlay) + 1 + xtraBytes + lImageSize + 20
            ReDim bExport(1 To lSizeArray)
            ' copy the compressed keys/tags to Export
            CopyMemory bExport(HeaderBytes + 17), bOverlay(0), UBound(bOverlay) + 1
            ' set pointer for start of image data
            lPointer = HeaderBytes + UBound(bOverlay) + 17
            bCompressMethod = (bCompressMethod Or 4)    ' update compression options used
            Erase bKeys     ' clear unneeded arrays
            Erase bOverlay
        Else
            lPointer = lPointer + 1 ' move from zero-bound to one-bound pointer
        End If
        
        ' place header information in the Export data
        CopyMemory bExport(HeaderBytes + 1), m_ItemCount, 4&                ' add nr items
        CopyMemory bExport(HeaderBytes + 5), m_KeyCount, 4&                 ' add nr keys
        CopyMemory bExport(HeaderBytes + 9), m_Width, 4&                    ' add width & height of images
        CopyMemory bExport(HeaderBytes + 13), m_Height, 4&
        
        ValidationOffset = lPointer     ' location where xtraBytes data can be found
        CopyMemory bExport(lPointer), xtraBytes, 4&
        ' Version 1:
            CopyMemory bExport(lPointer + 4), m_DelayLoad, 2&
        ' Version 2: add compression to bitmaps if cannot be converted to PNG
            CopyMemory bExport(lPointer + 6), bCompressMethod, 1&
        ' Version 3+: add extra bytes & adjust pointer appropriately
        '   ...
        ImageOffset = lPointer + xtraBytes  ' location where image data resides
        CopyMemory bExport(ImageOffset + 4), bImage(LBound(bImage)), lImageSize ' copy the image over
        xcom_ExportList = True
        
    End If
    
End Function

Friend Sub xcom_Clone(Action As Long)
    ' property page use only. Allows user to roll back if they cancel changes they made
    ' Parameter is a eCLoneAction enumeration
    
    If Action = cloneCreate Then
        If m_Clone Is Nothing Then         ' else already cloned; shouldn't be
            Set m_Clone = New cListImages  ' create new clone & tell it to clone from me
            If m_Clone.xcom_CloneData(Action, m_Keys(), m_Indexes(), m_Tags(), m_ItemCount, m_KeyCount, m_DelayLoad, m_Strip) = False Then
                Set m_Clone = Nothing
            End If
        End If
        
    ElseIf Action = cloneRestore Then       ' user rolled back changes, let's do it
        If Not m_Clone Is Nothing Then     ' else not cloned and we do nothing
            Dim Index As Long, pIndex As Long, modifiedCount As Long, bInIDE As Boolean
            ' tell clone to restore these parameters.
            modifiedCount = m_ItemCount
            If m_Clone.xcom_CloneData(Action, m_Keys(), m_Indexes(), m_Tags(), m_ItemCount, m_KeyCount, m_DelayLoad, m_Strip) Then
                ' need to rebuild our Images collection & pointer cross-references
                If Not modifiedCount = m_ItemCount Then
                    ReDim Preserve m_Images(1 To m_ItemCount)
                    For Index = modifiedCount + 1 To m_ItemCount      ' rebuild the collection, re-using existing dummy classes
                        Set m_Images(Index).ImageClass = New cListImage
                        m_Images(Index).ImageClass.xcom_InitializeItem Me
                    Next
                End If
                For Index = modifiedCount To 1 Step -1                ' clear any key references, all StrPtr values changed
                    m_Images(Index).Key = vbDefault
                Next
                For Index = 1 To m_KeyCount                           ' rebuild the key to pointer xRef, using previous m_Indexes array
                    pIndex = local_PointerToIndex(m_Keys(Index).ImagePtr)   ' get item index from pointer ref of this key
                    m_Keys(Index).ImagePtr = ObjPtr(m_Images(pIndex).ImageClass)
                    m_Images(pIndex).Key = StrPtr(m_Keys(Index).Key)
                Next
                For Index = 1 To m_ItemCount                          ' completely rebuild the pointer to index xRef
                    pIndex = local_FindPointer(ObjPtr(m_Images(Index).ImageClass), True, Index - 1)
                    If pIndex < Index Then
                        CopyMemory ByVal VarPtr(m_Indexes(pIndex + 1)), ByVal VarPtr(m_Indexes(pIndex)), (Index - pIndex) * 8&
                    End If
                    m_Indexes(pIndex).ImagePtr = ObjPtr(m_Images(Index).ImageClass)
                    m_Indexes(pIndex).Index = Index
                
                    m_Strip.gdiToken = mParent.ClassDataLong(eData_GDItoken, vbDefault)
                    Debug.Assert local_IsUnCompiled(bInIDE)
                    m_Strip.KeepGDIplusActive = Not bInIDE
                Next
            End If
            Set m_Clone = Nothing
        End If
        
    Else    ' cloneRelease :: destroy clone
        Set m_Clone = Nothing
    End If
End Sub

Friend Function xcom_CloneData(Action As Long, Keys() As KeyDataStructure, Indexes() As PointerStructure, Tags() As String, nrItems As Long, nrKeys As Long, bDelayMode As Boolean, cHost As c32bppDIB) As Boolean
    ' property page use only. Allows user to roll back if they cancel changes they made
    ' called within the clone's context. If getting this then this class is the clone
    
    If Action = cloneRestore Then           ' restore all image data
        Keys() = m_Keys()                   ' .ImagePtr member corrupted; needs rebuilding on return
        Indexes() = m_Indexes()             ' .ImagePtr member corrupted; needs rebuilding on return
        Tags() = m_Tags
        nrItems = m_ItemCount
        nrKeys = m_KeyCount
        bDelayMode = m_DelayLoad
        Set cHost = m_Strip                 ' restore imagelist
        Set m_Strip = Nothing               ' unref else Me.Clear erases the DIB
        xcom_CloneData = True
        
    ElseIf Action = cloneCreate Then
    
        Dim delayBytes() As Byte
        Set m_Strip = New c32bppDIB
        If cHost.GetOrginalFormat(delayBytes) Then
            m_Strip.SetOriginalFormat delayBytes    ' delay loaded; no image yet
            xcom_CloneData = True
        Else                                        ' copy the imagelist
            If m_Strip.InitializeDIB(cHost.Width, cHost.Height) Then
                CopyMemory ByVal m_Strip.BitsPointer, ByVal cHost.BitsPointer, m_Strip.scanWidth * m_Strip.Height
                xcom_CloneData = True
            End If
        End If
        If xcom_CloneData = True Then
            m_Keys() = Keys()               ' copy all important image data
            m_Indexes() = Indexes()
            m_Tags() = Tags()
            m_ItemCount = nrItems
            m_KeyCount = nrKeys
            m_DelayLoad = bDelayMode
        End If
    End If
End Function

Friend Property Get xcom_IsCloned() As Boolean
    ' property page use only.
    xcom_IsCloned = Not (m_Clone Is Nothing)
End Property

Friend Function xcom_AddItem_FromDIB(cHost As c32bppDIB, ByVal InsertAt As Long, ByVal Key As String, ByVal Tag As String, Scaled As Boolean, initImgCount As Long) As cListImage
    ' property page use only. Add an image or several images
    
    Select Case initImgCount
    Case Is > 1  ' bulk insert, append. Make room on the DIB and in arrays
        If iparseIsArrayEmpty(VarPtrArray(m_Images)) = vbDefault Then
            ReDim m_Images(1 To initImgCount + 2)
            ReDim m_Indexes(1 To initImgCount + 2)
            ReDim m_Tags(1 To initImgCount + 2)
        End If
        local_AdjustDibSize m_ItemCount + initImgCount, InsertAt, initImgCount  ' resize DIB if needed
        
    Case Is < vbDefault ' bulk insert with failures. When property page is trying to do bulk inserts
    '                     if an image fails to be loaded, property page skips it and continues on
    '                     but we have already shifted our DIB now we have to undo some of it
        local_AdjustDibSize m_ItemCount, InsertAt, initImgCount  ' resize DIB if needed
        
        
    Case vbDefault  ' initImgCount will be zero when bulk appending an image
        Set xcom_AddItem_FromDIB = local_AddItem(cHost, InsertAt, Key, Tag, Scaled, vbDefault)
    Case Else ' adding single image
        Set xcom_AddItem_FromDIB = local_AddItem(cHost, InsertAt, Key, Tag, Scaled, 1&)
    End Select
End Function

Friend Property Get xcom_ImagesDIB() As c32bppDIB
    ' property page use only
    Set xcom_ImagesDIB = m_Strip
End Property

Public Function xcom_IEnum() As IUnknown
Attribute xcom_IEnum.VB_UserMemId = -4
Attribute xcom_IEnum.VB_MemberFlags = "40"
    ' IMPORTANT: This property must be set to Procedure ID: -4, manually entered
    ' Also, check the Hide This Member box
    ' See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
    Set xcom_IEnum = CreateEnumerator(Me)
End Function


