VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cImageLists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' THE IMAGELISTS CLASS (cImageLists)
' Child class:  cImageList
'               ^^ child class: cListImages
'                               ^^ child class: cListImage
' - Contains collection of image lists and their keys
' - Creates and destroyes image lists
' - Returns specific image list classes
' - Provides Import/Export functions
' - See below for more

' Quick List of methods/properties
' ---------------------------------
'     Add creates a new image list
'     Clear destroys all images in the list
'     Compression gets/sets whether bitmap compression will be used when exporting image lists
'     CompressionNeeded returns whether Compression is recommended or not
'     Count returns number of images
'     ExportImageList exports an imagelist to an array suitable for files
'     ImageSizesFromStream returns number of icons, bit depths, color counts existing in passed array (usable for non-icons too)
'     ImportImageList imports an imagelist from array that was created with ExportImageList
'     IsKeyAssigned returns whether or not a passed Key is unique
'     Item returns a specific cImageList class
'     ManageGDIToken creates a GDI+ safe environment (called by usercontrol at startup)
'     MaxImages returns maximum number of images that are allowed on any image list
'     Remove deletes a specific image list
'     Version returns the image list class version

' Notes about indexing.
' 1. Referencing by Ordinal index is always faster than referencing by Key.
' 2. When you look at the cListImages class, it has 2 indexed arrays: Keys & Pointers
' This class only has the Key indexed array to get an ImageList by key. The Pointers
' array if used would be for backward communication from its child classes, however, since
' the amount of ImageLists will most likely be small (less than 50), adding the extra
' overhead of indexing pointers was opted against. Rather, iterating the collection is used.

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

' ***** THUNK APIs and VARIABLES (IDE GDI+ Usage, See ManageGDIToken)
Private Declare Function CreateWindowExA Lib "user32.dll" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function LoadLibrary Lib "kernel32.dll" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetParent Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function GetWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Const GWL_WNDPROC As Long = -4
Private Const GW_OWNER As Long = 4
Private Const WS_CHILD As Long = &H40000000

' ** duplicated from c32bppDIB suite which is not a public class, but this one is
Public Enum eILGrayScales
    lvil_gsclNTSCPAL = 0             ' R=R*.299, G=G*.587, B=B*.114 - Default
    lvil_gsclCCIR709 = 1             ' R=R*.213, G=G*.715, B=B*.072
    lvil_gsclSimpleAvg = 2           ' R,G, and B = (R+G+B)/3
    lvil_gsclRedMask = 3             ' uses only the Red sample value: RGB = Red / 3
    lvil_gsclGreenMask = 4           ' uses only the Green sample value: RGB = Green / 3
    lvil_gsclBlueMask = 5            ' uses only the Blue sample value: RGB = Blue / 3
    lvil_gsclRedGreenMask = 6        ' uses Red & Green sample value: RGB = (Red+Green) / 2
    lvil_gsclBlueGreenMask = 7       ' uses Blue & Green sample value: RGB = (Blue+Green) / 2
    lvil_gsclNone = -1
End Enum

' ** duplicated from c32bppDIB suite which is not a public class, but this one is
Public Enum eILConstants          ' See ImageSizesFromStream
    lvil_HIGH_COLOR = &HFFFF00
    lvil_TRUE_COLOR = &HFF000000
    lvil_TRUE_COLOR_ALPHA = &HFFFFFFFF
End Enum

Private Type KeyDataStructure
    Key As String
    ptrList As Long
End Type

Private m_Lists() As cImageList
Private m_Keys() As KeyDataStructure
Private m_ListCount As Long
Private m_GDItoken As Long
Private m_EnumIndex As Long

Private Const m_lvMagicNr As Long = &H4C56494C ' LVIL: LaVolpe ImageList
Private m_CompressOpt As Boolean        ' option to compress image (1) key/tag information (2)
Private Const m_Version As Long = 2
'^^ when new version implemented, must include new version code in cImageList.ExportImageList


' List of standard error messages used throughout image list classes
'    Err.Raise 1962, "ImageLists", "Invalid Key or Index"
'    Err.Raise 1963, "ImageLists", "Duplicated Key. Key already exists in this collection"
'
'    Err.Raise 1972, "ImageLists", "ImageList cannot support images of that size"
'    Err.Raise 1973, "ImageLists", "Maximum number of images have been exceeded. Create new image list"
'    Err.Raise 1974, "ImageLists", "Invalid ImageList Version. Update your ImageList classes/control"
'    Err.Raise 1975, "ImageLists", "Failed to Import external image list"
'    Err.Raise 1976, "ImageLists", "Error exporting image list. Ensure enough disk space and valid write-permissions"
'    Err.Raise 1977, "ImageLists", "Low system resource preventing appending additional images"

Implements IItemData


Public Property Get Item(Index As Variant) As cImageList
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"
    Set Item = local_KeyToObject(Index)
    ' IMPORTANT: This property must be set to Procedure ID: Default and also
    ' checked as User Interface Default. See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
End Property

' Returns/Sets option to compress image lists when exported or saved to the usercontrol.
' By default, if GDI+ or zLIB exists, imagelists are converted to PNG for storage. If
' this is not available, then this option determines if bitmap compression will be
' performed. Compression can reduce size by half but takes a little longer to
' load the image list. See CompressionNeeded property
Public Property Let Compression(useCompression As Boolean)
    m_CompressOpt = useCompression
End Property
Public Property Get Compression() As Boolean
    Compression = m_CompressOpt
End Property

Public Property Get CompressionNeeded() As Boolean
    ' returns if image list can be saved to PNG format or not
    ' If it can then property returns False, otherwise,
    ' property returns True. See CompressImageList
    Dim c32bpp As New c32bppDIB
    If c32bpp.isGDIplusEnabled = False Then
        CompressionNeeded = Not c32bpp.isZlibEnabled
    End If
End Property

Public Property Get Count() As Long         ' returns the number of image lists in the collection
    Count = m_ListCount
End Property
Public Property Get Version() As Long       ' returns the image list version.
    Version = m_Version                     ' Versions are not forward compatible, but are backward compatible
End Property

Public Sub Clear()                          ' destroys all image lists
    Erase m_Lists
    Erase m_Keys()
    m_ListCount = vbDefault
End Sub

Public Sub Remove(Index As Variant)         ' destroys a specific image list
    
    Dim kIndex As Long, iIndex As Long
    
    Call local_KeyToObject(Index, iIndex)
    If iIndex = vbDefault Then
        Err.Raise 1962, "ImageLists.Remove", "Invalid Key or Index"
        Exit Sub
    End If
    
    If m_ListCount = 1 Then
        Me.Clear                            ' last one
    Else
        ' need to shift the collection within the arrays
        Call local_PointerToKey(ObjPtr(m_Lists(iIndex)), kIndex)
        Set m_Lists(iIndex) = Nothing
        m_Keys(kIndex).Key = vbNullString   ' all imagelists have keys
        If iIndex < m_ListCount Then
            CopyMemory ByVal VarPtr(m_Lists(iIndex)), ByVal VarPtr(m_Lists(iIndex + 1)), (m_ListCount - iIndex) * 4&
            CopyMemory ByVal VarPtr(m_Lists(m_ListCount)), vbDefault, 4&
        End If
        If kIndex < m_ListCount Then
            CopyMemory ByVal VarPtr(m_Keys(kIndex)), ByVal VarPtr(m_Keys(kIndex + 1)), (m_ListCount - kIndex) * 8&
            FillMemory ByVal VarPtr(m_Keys(m_ListCount)), 8&, 0
        End If
        m_ListCount = m_ListCount - 1               ' decrement count
        ReDim Preserve m_Lists(1 To m_ListCount)    ' resize arrays
        ReDim Preserve m_Keys(1 To m_ListCount)
    End If

End Sub

Public Function MaxImages(ByVal imageWidth As Long, ByVal imageHeight As Long) As Long
    ' Image lists are a single bitmap that can contain 1,000s of individual images
    ' How large a single bitmap/DIB can be is dictated by hardware, not software.
    ' Using MSDN's recommendation, the following formula will be enforced
    If imageWidth > vbDefault Then
        If imageHeight > vbDefault Then MaxImages = (1024000 \ (imageWidth * imageHeight))
    End If
End Function

Public Function Add(ByVal Width As Long, ByVal Height As Long, _
                ByVal Key As String, Optional ByVal InitImageCount As Long = 1&) As cImageList
                
    ' Function creates a new image list
    ' Parameters:
    '   Width :: the width every image will be resized to when added to the image list
    '   Height :: the height every image will be resized to when added to the image list
    '   Key :: required Key for the image list. Keys can be numeric, cannot be null
    '   InitImageCount :: initialization image count for the imagelist, 1 is always default
    '       pre-sizing imagelist can speed up adding multiple images
                
    Dim newList As cImageList
    Dim kIndex As Long, bUnique As Boolean
    
    If Not Key = vbNullString Then kIndex = local_FindKey(Key, bUnique, m_ListCount)
    If Not bUnique Then
        Err.Raise 1963, "ImageLists.Add", "Duplicated Key. Key already exists in this collection"
        Exit Function
    End If
    
    If InitImageCount < 1& Then InitImageCount = 1&     ' ensure initial size never less than one
    ' ensure image list can support the requested size and number of images
    If MaxImages(Width, Height) < InitImageCount Then
        Err.Raise 1973, "ImageLists.Add", "Maximum number of images have been exceeded. Create new image list"
        Exit Function
    End If
    
    m_ListCount = m_ListCount + 1               ' increment count
    ReDim Preserve m_Lists(1 To m_ListCount)    ' resize arrays
    ReDim Preserve m_Keys(1 To m_ListCount)
    
    Set newList = New cImageList                ' create new list
    Set m_Lists(m_ListCount) = newList
    
    local_AddKey Key, kIndex, m_ListCount       ' add the key to the keys collection
    CopyMemory ByVal VarPtr(Key), vbDefault, 4& ' and cross-reference the key
    
    ' initialize the new list
    Call newList.xcom_InitializeList(Width, Height, InitImageCount, Me)
    Set Add = newList
    
End Function
                
Public Function IsKeyAssigned(Key As String) As Long
    ' Returns whether or not passed Key has been assigned
    ' Return value of 0=unique key, else is the Index of the image having that key
    Dim kIndex As Long, bUnique As Boolean
    kIndex = local_FindKey(Key, bUnique, m_ListCount)
    If Not bUnique Then IsKeyAssigned = local_PointerToIndex(m_Keys(kIndex).ptrList)
End Function

Public Function ImportImageList(inData() As Byte, Optional Key As String) As cImageList

    ' Function parses the inData() to create an image list
    ' Parameters:
    '   inData() :: an array of bytes that comprise a formatted imagelist
    '       When an image list is exported to file, simply read the file bytes and pass them here
    '   Key [out] :: the key assigned to the imported image list
    
    Dim ValidationOffset As Long, ImageOffset As Long
    Dim vVersion As Long, vData As Long, lPointer As Long
    Dim lBoundOffset As Long
    
    ' validation first
    If iparseArrayProps(VarPtrArray(inData), vData) = 0& Then Exit Function
    If vData > 1& Then Exit Function
    
    ' validate the array is our imagelist
    lPointer = LBound(inData)
    lBoundOffset = lPointer - 1
    If Abs(UBound(inData) - LBound(inData)) < 20 Then Exit Function   ' minimal length need to validate
    
    CopyMemory vData, inData(lPointer), 4&                  ' get magic number which is LVIL
    If vData <> m_lvMagicNr Then Exit Function
        
    CopyMemory ValidationOffset, inData(lPointer + 8), 4&   ' location to IL's xtraData offset
    CopyMemory ImageOffset, inData(lPointer + 12), 4&       ' location to IL's image offset
    If UBound(inData) < (ImageOffset - lPointer + 1) Then Exit Function ' invalid data
    
    ' get the number of xtraBytes used by the IL version & add it to the Validation offset
    CopyMemory vData, inData(ValidationOffset + lBoundOffset), 4&
    ' if the ImageOffset does not equal the sum, then invalid data
    If ImageOffset <> (vData + ValidationOffset) Then Exit Function
    
    Dim keyLen As Long, baseKey As String
    Dim kIndex As Long, bOK As Boolean, fIndex As Long
    Dim tList As cImageList
    
    ' appears to be valid data. Get its version
    CopyMemory vVersion, inData(lPointer + 4), 4&
    lPointer = lPointer + 16                                ' move to imagelist name
    If vVersion > m_Version Then                            ' version > our version, so we can't import future version data
        Err.Raise 1974, "ImageLists.ImportImageList", "Invalid ImageList Version. Update your ImageList classes/control"
        Exit Function
    End If
    
    m_ListCount = m_ListCount + 1                           ' create a new structure for the imagelist
    If iparseIsArrayEmpty(VarPtrArray(m_Lists)) = 0 Then
        ReDim Preserve m_Lists(1 To m_ListCount)            ' initialize lists
        ReDim Preserve m_Keys(1 To m_ListCount)
    ElseIf m_ListCount > UBound(m_Lists) Then               ' increment arrays
        ReDim Preserve m_Lists(1 To m_ListCount + 1)
        ReDim Preserve m_Keys(1 To m_ListCount + 1)
    End If
    
    CopyMemory keyLen, inData(lPointer), 4&                     ' length of imagelist key
    Key = Space$(keyLen \ 2)                                    ' create string neeeded
    CopyMemory ByVal StrPtr(Key), inData(lPointer + 4), keyLen  ' copy the key
    
    Set tList = New cImageList                                  ' create a new imagelist
    tList.xcom_InitializeList vbDefault, vbDefault, 0&, Me      ' initialize it (we won't know width/height yet)
    Set m_Lists(m_ListCount) = tList
    
    If m_ListCount = 1 Then                                     ' handle 1st occurence specifically
        kIndex = 1
        m_Keys(kIndex).Key = Key                                ' set the key & cross-reference
        m_Keys(kIndex).ptrList = ObjPtr(tList)                  ' cross-reference the key
    Else
        kIndex = local_FindKey(Key, bOK, m_ListCount - 1)       ' find insert position of key in collection
        If Not bOK Then                                         ' auto-create unique key if needed
            baseKey = Key
            Do Until bOK
                fIndex = fIndex + 1
                Key = baseKey & "(" & fIndex & ")"
                kIndex = local_FindKey(Key, bOK, m_ListCount - 1) ' find insert position of key in array
            Loop
        End If
        baseKey = Key
        local_AddKey baseKey, kIndex, m_ListCount               ' insert the key into the array
        CopyMemory ByVal VarPtr(baseKey), vbDefault, 4&         ' zeroize the passed key
        bOK = False
    End If
    
    On Error Resume Next
    ' import the image data into the imagelist
    bOK = tList.Images.xcom_ImportList(inData(), lPointer + keyLen + 3 - lBoundOffset, ValidationOffset, vVersion)
    If bOK Then
        ' when successful, the cImageList class called back to this class to update the imagelist width & height properties
        Set ImportImageList = tList
    Else
        ' failed to import, rollback
        m_Keys(kIndex).Key = vbNullString
        m_Keys(kIndex).ptrList = vbDefault
        Set m_Lists(m_ListCount) = Nothing
        m_ListCount = m_ListCount - 1
        If Err Then
            Err.Raise Err.Number, "ImageLists.ImportImageList", Err.Description
            Err.Clear
        Else
            Err.Raise 1975, "ImageLists.ImportImageList", "Failed to Import external image list"
        End If
    End If

End Function

Public Function ExportImageList(Index As Variant, outData() As Byte) As Boolean

    ' Function exports an imagelist, its key, images, images' keys/tags to an array
    ' Parameters:
    '   Index :: Ordinal or Key of the imagelist to be exported
    '   outData() :: an array which will be resized on return
    ' Return Value: If false, the outData array should be considered invalid
    
    Dim sKey As String, kIndex As Long
    Dim lPointer As Long, buffLen As Long
    Dim ValidationOffset As Long, ImageOffset As Long
    Dim ilExport As cImageList
    Dim bOK As Boolean
    
    ' validate passed index
    Set ilExport = local_KeyToObject(Index)
    If ilExport Is Nothing Then
        Err.Raise 1962, "ImageLists.ExportImageList", "Invalid Key or Index"
        Exit Function
    End If
        
    sKey = local_PointerToKey(ObjPtr(ilExport))     ' get image list's key & key size
    buffLen = Len(sKey) * 2 + 20                    ' add header size too
    
    On Error Resume Next
    ' call child function to export images, keys & tags to our buffer
    bOK = ilExport.Images.xcom_ExportList(outData, buffLen, ValidationOffset, ImageOffset, m_CompressOpt)
    If bOK = True Then
        lPointer = LBound(outData)                              ' add our data
        CopyMemory outData(lPointer), m_lvMagicNr, 4&           ' add our magic number
        ' next 4 bytes are for version. 0 indicates version 1
        CopyMemory outData(lPointer + 4), m_Version, 4&         ' version
        CopyMemory outData(lPointer + 8), ValidationOffset, 4&  ' used to validate array for import
        CopyMemory outData(lPointer + 12), ImageOffset, 4&      ' position of image in array relative to byte 1
        buffLen = buffLen - 20                                  ' remove header size
        CopyMemory outData(lPointer + 16), buffLen, 4&          ' store the length & the key
        CopyMemory outData(lPointer + 20), ByVal StrPtr(sKey), buffLen
        ExportImageList = True
    Else
        If Err Then
            Err.Raise Err.Number, "ImageLists.ExportImageList", Err.Description
        Else
            Err.Raise 1976, "ImageLists.ExportImageList", "Error exporting image list. Ensure enough disk space and valid write-permissions"
        End If
    End If

End Function

Public Function ImageSizesFromStream(inStream() As Byte, arraySizes() As Long) As Long

    ' Function will return the number of images, bit depths and color count contained
    ' within the passed array.  For non multi-resource files, the return value would be one.
    ' If the passed array contains no recognized image, then the return value is zero

    ' arraySizes() will be populated as follows
    ' 1st dimension of the array
    '   0 element: image/icon width
    '   1 element: image/icon height
    '   The next two elements only apply if array contains multiple images/icons
    '   2 element: image/icon bit depth: 1,2,4,8,16,24,32
    '   3 element: image/icon colors: 2,4,16,256,HIGH_COLOR,TRUE_COLOR, TRUE_COLOR_ALPHA
    '   -- HIGH_COLOR, TRUE_COLOR,& TRUE_COLOR_ALPHA are public variables
    ' 2nd dimension of the array is 0 to number of icons - 1
    
    Dim tImage As New c32bppDIB, nrImages As Long
    If tImage.LoadPicture_Stream(inStream, , , , , True) = True Then
        nrImages = tImage.SourceIconSizes(arraySizes)
        If nrImages = 0& Then
            nrImages = 1
            ReDim arraySizes(0 To 3, 0 To 0)
            arraySizes(0, 0) = tImage.Width
            arraySizes(1, 0) = tImage.Height
        End If
        Set tImage = Nothing
    Else
        ReDim arraySizes(0 To 3, -1 To -1)
    End If
    ImageSizesFromStream = nrImages

End Function
                
                
' LOCAL-USE FUNCTIONS ONLY

Private Sub local_AddKey(Key As String, Index As Long, lstIndex As Long)
    ' adds a key to the key collection ensuring ascending sort order
    ' The passed Index indicates the sorted position for the new key
    If Index < m_ListCount Then
        CopyMemory ByVal VarPtr(m_Keys(Index + 1)), ByVal VarPtr(m_Keys(Index)), (m_ListCount - Index) * 8&
    End If
    CopyMemory ByVal VarPtr(m_Keys(Index).Key), StrPtr(Key), 4& ' copies key pointer; key must be zeroized by calling function
    m_Keys(Index).ptrList = ObjPtr(m_Lists(lstIndex))           ' cross-reference key to class pointer
    
End Sub

Private Function local_KeyToObject(Key As Variant, Optional itemIndex As Long) As cImageList
    ' converts an object pointer to a cImageList class and also returns the class index
    Dim kIndex As Long, bIsNew As Boolean
    
    Select Case VarType(Key)
    Case vbString                           ' Strings are always treated as Keys
        If Not Key = vbNullString Then
            kIndex = local_FindKey(CStr(Key), bIsNew, m_ListCount)
            If bIsNew = False Then
                itemIndex = local_PointerToIndex(m_Keys(kIndex).ptrList)
                Set local_KeyToObject = m_Lists(itemIndex)
            End If
        End If
    Case vbLong, vbInteger, vbDouble, vbByte, vbSingle ' Numerical varTypes are treated as Indexes
        On Error Resume Next
        kIndex = Int(Key)
        If Err Then
            Err.Clear
        Else
            If kIndex > vbDefault And kIndex <= m_ListCount Then
                itemIndex = kIndex
                Set local_KeyToObject = m_Lists(kIndex)
            End If
        End If
    End Select

End Function

Private Function local_FindKey(Key As String, ByRef isNew As Boolean, keyCount As Long) As Long

    ' BINARY SEARCH ALGORITHM -- Divide and conquer.
    ' Binary search algorithms are about the fastest on the planet, but
    ' its biggest disadvantage is that the array must already be sorted.
    ' Ex: binary search can find a value among 1 million values between just 1 and 20 iterations
    
    ' [in] Key. A value to search for. Order is always ascending
    ' [out] isNew. If Key not found, isNew is True else False
    ' [out] Return value: The Index where Key was found or where the new Key should be inserted

    Dim UB As Long, LB As Long
    Dim newIndex As Long
    
    If keyCount = 0& Then
        newIndex = 1&
        isNew = True
    Else
    
        UB = keyCount
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
                
Private Sub local_ReKeyList(Index As Long, newKey As String)
    ' function changes an image list's key
    ' Index parameter validated before passed here
    If newKey = vbNullString Then
        Err.Raise 1962, "ImageLists.Key", "Invalid Key"
        Exit Sub
    End If
    
    Dim kIndex As Long, newKindex As Long
    Dim bUnique As Boolean
    
    Call local_PointerToKey(ObjPtr(m_Lists(Index)), kIndex) ' get key index from class index
    newKindex = local_FindKey(newKey, bUnique, m_ListCount) ' see if new key duplicates a key
    
    If Not bUnique Then
        If kIndex = newKindex Then                          ' key duplicated but for same class
            m_Keys(kIndex).Key = newKey                     ' make change; may be a case-sensitive change only
        Else
            Err.Raise 1963, "ImageLists.Key", "Duplicated Key. Key already exists in this collection"
        End If
        Exit Sub
    End If
    
    ' Key is changing...
    m_Keys(kIndex).Key = vbNullString                       ' remove old key
    If kIndex < m_ListCount Then                            ' shift key collection as needed
        CopyMemory ByVal VarPtr(m_Keys(kIndex)), ByVal VarPtr(m_Keys(kIndex + 1)), (m_ListCount - kIndex) * 8&
        CopyMemory ByVal VarPtr(m_Keys(m_ListCount).Key), vbDefault, 4&
    End If
    kIndex = local_FindKey(newKey, True, m_ListCount - 1)   ' get position in shifted array for new key
    local_AddKey newKey, kIndex, Index                      ' add the key to the collection
    CopyMemory ByVal VarPtr(newKey), vbDefault, 4&          ' zeroize the new key
    
End Sub

Private Sub local_ReIndexLists(oldIndex As Long, newIndex As Long)
    ' function changes the Index of a imagelist
    ' oldIndex validated before passed here
    If oldIndex = newIndex Then Exit Sub
    If newIndex > m_ListCount Or newIndex < 1 Then
        Err.Raise 1962, "ImageLists.Index", "Invalid Index"
        Exit Sub
    End If
    
    Dim oldPtr As Long
    CopyMemory oldPtr, ByVal VarPtr(m_Lists(oldIndex)), 4&  ' save pointer of class in oldIndex
    If newIndex < oldIndex Then ' moving imagelist higher in index order
        CopyMemory ByVal VarPtr(m_Lists(newIndex + 1)), ByVal VarPtr(m_Lists(newIndex)), (oldIndex - newIndex) * 4&
    Else                        ' moving imagelist lower in index order
        CopyMemory ByVal VarPtr(m_Lists(oldIndex)), ByVal VarPtr(m_Lists(oldIndex + 1)), (newIndex - oldIndex) * 4&
    End If
    CopyMemory ByVal VarPtr(m_Lists(newIndex)), oldPtr, 4&  ' restore pointer
    
End Sub
                
Private Function local_PointerToIndex(oPtr As Long) As Long
    ' function returns a class Index from its pointer
    ' classes are not indexed since the number of them are not expected to be very large
    Dim pIndex As Long
    For pIndex = 1 To m_ListCount
        If ObjPtr(m_Lists(pIndex)) = oPtr Then Exit For
    Next
    local_PointerToIndex = pIndex
End Function

Private Function local_PointerToKey(oPtr As Long, Optional keyIndex As Long) As String
    ' function returns a Key from a class pointer
    ' classes are not indexed since the number of them are not expected to be very large
    For keyIndex = m_ListCount To 1 Step -1
        If m_Keys(keyIndex).ptrList = oPtr Then Exit For
    Next
    local_PointerToKey = m_Keys(keyIndex).Key
End Function
                

' CLASS CROSS-COMMUNICATION USE ONLY. Allows child classes to get/set parent properties

Private Property Let IItemData_ClassDataLong(DataName As eDataNames, oID As Long, ByVal newValue As Long)
    Dim iIndex As Long
    Select Case DataName
        Case eData_Index:
            iIndex = local_PointerToIndex(oID)
            If iIndex Then local_ReIndexLists iIndex, newValue
        Case eData_EnumReset: m_EnumIndex = vbDefault
    End Select
End Property
Private Property Get IItemData_ClassDataLong(DataName As eDataNames, oID As Long) As Long
    Dim iIndex As Long
    Select Case DataName
        Case eData_Index: IItemData_ClassDataLong = local_PointerToIndex(oID)
        Case eData_GDItoken: IItemData_ClassDataLong = m_GDItoken
        Case eData_MaxImgCount:
            iIndex = local_PointerToIndex(oID)
            If iIndex Then IItemData_ClassDataLong = MaxImages(m_Lists(iIndex).Width, m_Lists(iIndex).Height)
        Case eData_EnumNext
            m_EnumIndex = m_EnumIndex + 1
            IItemData_ClassDataLong = (m_EnumIndex <= m_ListCount)
    End Select
End Property

Private Property Get IItemData_ClassDataObject(DataName As eDataNames, oID As Long) As Variant
    ' Note: Kelly Ethridge example uses VariantCopy API to transfer enumerated item
    ' to property; however, no can do with objects it appears -- crash. This works.
    If DataName = eData_EnumItem Then Set IItemData_ClassDataObject = m_Lists(m_EnumIndex)
End Property

Private Property Let IItemData_ClassDataString(DataName As eDataNames, oID As Long, ByVal newValue As String)
    Dim iIndex As Long
    Select Case DataName
        Case eData_Key:
            iIndex = local_PointerToIndex(oID)
            If iIndex Then local_ReKeyList iIndex, newValue
    End Select
End Property
Private Property Get IItemData_ClassDataString(DataName As eDataNames, oID As Long) As String
    Select Case DataName
        Case eData_Key: IItemData_ClassDataString = local_PointerToKey(oID)
    End Select
End Property


Private Sub Class_Initialize()
    m_CompressOpt = True
End Sub

Private Sub Class_Terminate()
    Me.Clear
End Sub

' Note: If you remove the usercontrol from your project,
' this routine should be called prior to creating the first image list

Friend Function ManageGDIToken(ByVal hwnd As Long) As Long
        
    ' Parameter: hWnd is any valid hWnd within your project
        
    ' Return value is an hWnd to the GDI+ IDE-safe monitor.
    ' If return value is zero, then GDI+ is not installed on the O/S
        
    ' So this function will create (if needed) only one API window for the entire
    ' VB instance and remains as long as VB is opened. The API window is a child
    ' of a hidden VB top-level window and when that window closes, so will the
    ' API window which then will release GDI+. By keeping its own reference to
    ' the GDI+ library, there is no way for that library to be unloaded without
    ' the API window closing.
    ' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69434&lngWId=1
        
    Dim hwndGDIsafe     As Long                 'API window to monitor IDE shutdown
    Dim bOK As Boolean
    Dim cGDIp As cGDIPlus
    
    ' find the hidden VB owner window. All VB apps have this (run-time & design-time).
    If hwnd = 0& Then Exit Function
    Do
        hwndGDIsafe = GetParent(hwnd)
        If Not hwndGDIsafe = 0& Then hwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    ' ok, got the highest level parent, now find highest level owner
    Do
        hwndGDIsafe = GetWindow(hwnd, GW_OWNER)
        If Not hwndGDIsafe = 0& Then hwnd = hwndGDIsafe
    Loop Until hwndGDIsafe = 0&
    
    hwndGDIsafe = FindWindowEx(hwnd, 0&, "Static", "GDI+Safe Patch")
    If hwndGDIsafe Then
        ManageGDIToken = hwndGDIsafe    ' we already have a manager running for this VB instance
        m_GDItoken = True
        Exit Function                   ' can abort
    End If
    
    On Error Resume Next
    Set cGDIp = New cGDIPlus
    If cGDIp.isGDIplusOk(m_GDItoken, True) = False Then Exit Function
    If Err Then
        m_GDItoken = 0&
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    Dim z_ScMem         As Long                 'Thunk base address
    Dim z_Code()        As Long                 'Thunk machine-code initialised here
    Dim nAddr           As Long                 'hwndGDIsafe prev window procedure

    Const WNDPROC_OFF   As Long = &H30          'Offset where window proc starts from z_ScMem
    Const PAGE_RWX      As Long = &H40&         'Allocate executable memory
    Const MEM_COMMIT    As Long = &H1000&       'Commit allocated memory
    Const MEM_RELEASE   As Long = &H8000&       'Release allocated memory flag
    Const MEM_LEN       As Long = &HD4          'Byte length of thunk
        
    z_ScMem = VirtualAlloc(0, MEM_LEN, MEM_COMMIT, PAGE_RWX) 'Allocate executable memory
    If z_ScMem <> 0 Then                                     'Ensure the allocation succeeded
        
        ' we make the api window a child so we can use FindWindowEx to locate it easily
        hwndGDIsafe = CreateWindowExA(0&, "Static", "GDI+Safe Patch", WS_CHILD, 0&, 0&, 0&, 0&, hwnd, 0&, App.hInstance, ByVal 0&)
        If hwndGDIsafe <> 0 Then
        
            ReDim z_Code(0 To MEM_LEN \ 4 - 1)
        
            z_Code(12) = &HD231C031: z_Code(13) = &HBBE58960: z_Code(14) = &H12345678: z_Code(15) = &H3FFF631: z_Code(16) = &H74247539: z_Code(17) = &H3075FF5B: z_Code(18) = &HFF2C75FF: z_Code(19) = &H75FF2875
            z_Code(20) = &H2C73FF24: z_Code(21) = &H890853FF: z_Code(22) = &HBFF1C45: z_Code(23) = &H2287D81: z_Code(24) = &H75000000: z_Code(25) = &H443C707: z_Code(26) = &H2&: z_Code(27) = &H2C753339: z_Code(28) = &H2047B81: z_Code(29) = &H75000000
            z_Code(30) = &H2C73FF23: z_Code(31) = &HFFFFFC68: z_Code(32) = &H2475FFFF: z_Code(33) = &H681C53FF: z_Code(34) = &H12345678: z_Code(35) = &H3268&: z_Code(36) = &HFF565600: z_Code(37) = &H43892053: z_Code(38) = &H90909020: z_Code(39) = &H10C261
            z_Code(40) = &H562073FF: z_Code(41) = &HFF2453FF: z_Code(42) = &H53FF1473: z_Code(43) = &H2873FF18: z_Code(44) = &H581053FF: z_Code(45) = &H89285D89: z_Code(46) = &H45C72C75: z_Code(47) = &H800030: z_Code(48) = &H20458B00: z_Code(49) = &H89145D89
            z_Code(50) = &H81612445: z_Code(51) = &H4C4&: z_Code(52) = &HC63FF00

            z_Code(1) = 0                                                   ' shutDown mode; used internally by ASM
            z_Code(2) = local_zFnAddr("user32", "CallWindowProcA")          ' function pointer CallWindowProc
            z_Code(3) = local_zFnAddr("kernel32", "VirtualFree")            ' function pointer VirtualFree
            z_Code(4) = local_zFnAddr("kernel32", "FreeLibrary")            ' function pointer FreeLibrary
            z_Code(5) = m_GDItoken                                          ' Gdi+ token
            z_Code(10) = LoadLibrary("gdiplus")                             ' library pointer (add reference)
            z_Code(6) = GetProcAddress(z_Code(10), "GdiplusShutdown")       ' function pointer GdiplusShutdown
            z_Code(7) = local_zFnAddr("user32", "SetWindowLongA")           ' function pointer SetWindowLong
            z_Code(8) = local_zFnAddr("user32", "SetTimer")                 ' function pointer SetTimer
            z_Code(9) = local_zFnAddr("user32", "KillTimer")                ' function pointer KillTimer
        
            z_Code(14) = z_ScMem                                            ' ASM ebx start point
            z_Code(34) = z_ScMem + WNDPROC_OFF                              ' subclass window procedure location
        
            RtlMoveMemory z_ScMem, VarPtr(z_Code(0)), MEM_LEN               'Copy the thunk code/data to the allocated memory
        
            nAddr = SetWindowLong(hwndGDIsafe, GWL_WNDPROC, z_ScMem + WNDPROC_OFF) 'Subclass our API window
            RtlMoveMemory z_ScMem + 44, VarPtr(nAddr), 4& ' Add prev window procedure to the thunk
            
            bOK = True
            ManageGDIToken = hwndGDIsafe    ' return handle of our GDI+ manager
        
        Else
        
            VirtualFree z_ScMem, 0, MEM_RELEASE     ' failure - release memory
            z_ScMem = 0&
            'MsgBox "Failed to create monitoring window", vbExclamation + vbOKOnly, "Error"
            
        End If
    
    Else
        
      VirtualFree z_ScMem, 0, MEM_RELEASE           ' failure - release memory
      z_ScMem = 0&
      'MsgBox "Failed to initialize thunk memory", vbExclamation + vbOKOnly, "Error"
        
    End If
    
    If bOK = False Then
        Call cGDIp.InitializeGDIplus(m_GDItoken, True) ' release token if error occurred
        m_GDItoken = 0&
    End If
    
End Function


Private Function local_zFnAddr(ByVal sDLL As String, ByVal sProc As String) As Long
'Return the address of the specified DLL/procedure

  local_zFnAddr = GetProcAddress(GetModuleHandleA(sDLL), sProc)  'Get the specified procedure address
  Debug.Assert local_zFnAddr                                     'In the IDE, validate that the procedure address was located
  ' ^^ FYI VB5 users. Search for local_zFnAddr("vba6", "EbMode") and replace with local_zFnAddr("vba5", "EbMode")

End Function

Public Function xcom_IEnum() As IUnknown
Attribute xcom_IEnum.VB_UserMemId = -4
Attribute xcom_IEnum.VB_MemberFlags = "40"
    ' IMPORTANT: This property must be set to Procedure ID: -4, manually entered
    ' Also, check the Hide This Member box
    ' See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
    Set xcom_IEnum = CreateEnumerator(Me)
End Function

