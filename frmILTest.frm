VERSION 5.00
Object = "*\ALaVolpeImageList.vbp"
Begin VB.Form frmILTest 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   ScaleHeight     =   5325
   ScaleWidth      =   6555
   StartUpPosition =   3  'Windows Default
   Begin LaVolpeImageList.lvImageList lvImageList1 
      Left            =   1320
      Top             =   1260
      _ExtentX        =   1058
      _ExtentY        =   1058
      Img1            =   "frmILTest.frx":0000
      Img2            =   "frmILTest.frx":C0B9
      Count           =   2
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Do Enumeration Test"
      Height          =   405
      Left            =   3375
      TabIndex        =   12
      Top             =   3135
      Width           =   2490
   End
   Begin VB.CheckBox chkExtractEx 
      Caption         =   "Extract As Rendered"
      Height          =   285
      Left            =   3390
      TabIndex        =   11
      Top             =   4185
      Width           =   2040
   End
   Begin VB.TextBox txtKey 
      Height          =   345
      Left            =   1020
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4575
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   945
      Left            =   105
      ScaleHeight     =   885
      ScaleWidth      =   795
      TabIndex        =   9
      Top             =   4020
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Extract Icon Example"
      Height          =   495
      Index           =   0
      Left            =   1020
      TabIndex        =   8
      Top             =   4065
      Width           =   2040
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "Mirror Horizontal"
      Height          =   285
      Index           =   4
      Left            =   3390
      TabIndex        =   7
      Top             =   1920
      Width           =   2070
   End
   Begin VB.OptionButton optSize 
      Caption         =   "48 x 48"
      Height          =   270
      Index           =   1
      Left            =   4560
      TabIndex        =   0
      Top             =   2280
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.OptionButton optSize 
      Caption         =   "32 x 32"
      Height          =   270
      Index           =   0
      Left            =   3375
      TabIndex        =   6
      Top             =   2265
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   3825
      Left            =   75
      Picture         =   "frmILTest.frx":1792B
      ScaleHeight     =   251
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   204
      TabIndex        =   1
      Top             =   165
      Width           =   3120
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "60% transparent"
      Height          =   285
      Index           =   0
      Left            =   3345
      TabIndex        =   2
      Top             =   480
      Width           =   2070
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "90 degree rotation clockwisse"
      Height          =   285
      Index           =   1
      Left            =   3360
      TabIndex        =   3
      Top             =   840
      Width           =   2490
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "5% darker"
      Height          =   285
      Index           =   2
      Left            =   3375
      TabIndex        =   4
      Top             =   1200
      Width           =   2070
   End
   Begin VB.CheckBox chkTest 
      Caption         =   "Gray Scaled"
      Height          =   285
      Index           =   3
      Left            =   3390
      TabIndex        =   5
      Top             =   1560
      Width           =   2070
   End
   Begin VB.Label Label1 
      Caption         =   "^^ Note: Windows 2000 and lower will display most 32bpp icons very poorly."
      Height          =   300
      Left            =   120
      TabIndex        =   13
      Top             =   5025
      Width           =   5940
   End
End
Attribute VB_Name = "frmILTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' for some additional info, jump to bottom of this module

Private Declare Function DestroyIcon Lib "user32.dll" (ByVal hIcon As Long) As Long
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long

Private Sub DoSample()

    Dim X As Long, y As Long, Index As Long
    Dim imgCx As Long, imgCy As Long, xOffset As Long
    Dim mirrorWidth As Long
    Dim Key As String
    
    Dim lDarker As Long
    Dim lRotation As Long
    Dim lOpacity As Long
    Dim lGrayScale As Long
    
    If chkTest(0) Then lOpacity = 40 Else lOpacity = 100
    If chkTest(1) Then lRotation = 90
    If chkTest(2) Then lDarker = -5
    If chkTest(3) Then lGrayScale = lvil_gsclCCIR709 Else lGrayScale = lvil_gsclNone
    
    Picture1.Cls
    If optSize(0) Then Key = "ImgLst32" Else Key = "ImgLst48"
    With lvImageList1.ImageLists(Key)
        imgCx = .Width
        imgCy = .Height
        xOffset = (Picture1.ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, vbPixels) Mod (imgCx + imgCx \ 4)) \ 2
        X = xOffset
        y = 5
        
        If chkTest(4) Then mirrorWidth = -imgCx Else mirrorWidth = imgCx
        
        For Index = 1 To .Images.Count
            .Images.Render Index, Picture1.hdc, X, y, mirrorWidth, imgCy, , , , , lOpacity, lGrayScale, lDarker, lRotation
            X = X + imgCx + imgCx \ 4
            If X + imgCx > Picture1.ScaleWidth Then
                X = xOffset
                y = y + imgCx + 8
            End If
        Next
    End With
    Picture1.Refresh

End Sub

Private Sub chkExtractEx_Click()
    Call Command1_Click(-1)
End Sub

Private Sub Command1_Click(Index As Integer)

    ' Extracting image as icon example.
    ' The icon returned will be the same size as the source imagelist item.
    ' The bit depth of the icon can be anywhere between 1, 4, 8, 24 & 32
    ' As a general rule of thumb:
    '   - rotated images are 32bpp
    '   - stretched/resized images are 32bpp if resized with bicubic or bilinear interpolation
    '   - non-resized images are original bit depth
    
    ' There are 2 functions provided by the image list class
    ' ExtractIcon.  Returns the image list image as is, as an icon
    ' ExtractIconEx. Returns the image list image rendered with options, as an icon
    ' Both functions have ability to also optionally return the icon's width, height & bit depth
    
    Dim hIcon As Long
    Dim Key As String
    Dim imgIndex As Long
    
    Dim iconCx As Long, iconCy As Long
    '^^ the 2 extract icon functions can return icon's width/height
    ' Unless rotating the image, the width/height will be the same as the image list
    ' But when rotating the width height can be = Sqr(Width^2 + Height^2)
    ' Nice to know info should you want to center an image on a point
    
    Dim lDarker As Long
    Dim lRotation As Long
    Dim lOpacity As Long
    Dim lGrayScale As Long
    Dim bMirrored As Boolean
    
    If optSize(0) Then Key = "ImgLst32" Else Key = "ImgLst48"
    If Index = 0 Then
        imgIndex = Val(InputBox("Enter the image Index to sample. Index 1 to " & lvImageList1.ImageLists(Key).Images.Count, "Extract Icon Sample", Command1(0).Tag))
    Else
        imgIndex = Val(Command1(0).Tag)
    End If
    
    If imgIndex Then
        On Error GoTo EH
        
        If chkExtractEx.Value = 1 Then  ' sample ExtractIconEx call
            If chkTest(0) Then lOpacity = 40 Else lOpacity = 100
            If chkTest(1) Then lRotation = 90
            If chkTest(2) Then lDarker = -5
            If chkTest(3) Then lGrayScale = lvil_gsclCCIR709 Else lGrayScale = lvil_gsclNone
            If chkTest(4) Then bMirrored = True
            hIcon = lvImageList1.ImageLists(Key).Images.ExtractIconEx(imgIndex, bMirrored, , lOpacity, lGrayScale, lDarker, lRotation, iconCx, iconCy)
            
        Else                            ' sample ExtractIcon call
            hIcon = lvImageList1.ImageLists(Key).Images.ExtractIcon(imgIndex)
            iconCx = lvImageList1.ImageLists(Key).Width
            iconCy = lvImageList1.ImageLists(Key).Height
            
        End If
        
        Command1(0).Tag = imgIndex
        If hIcon Then
            Picture2.Cls
            DrawIconEx Picture2.hdc, (Picture2.ScaleWidth - iconCx) \ 2, (Picture2.ScaleHeight - iconCy) \ 2, hIcon, 0, 0, 0, 0, &H3
            txtKey.Text = "Key: " & lvImageList1.ImageLists(Key).Images(imgIndex).Key
            DestroyIcon hIcon
        End If
    
    End If
    
EH:
If Err Then
    MsgBox Err.Description, vbInformation + vbOKOnly
    Err.Clear
End If

End Sub

Private Sub Command2_Click()
    Call EnumerationExample
    MsgBox "Open your Debug/Immediate window", vbOKOnly + vbInformation, "Done"
End Sub

Private Sub Form_Load()
    Picture2.ScaleMode = vbPixels
    Picture2.Picture = Picture1.Picture
    Command1(0).Tag = 1    ' see command1's Inputbox
    optSize(0).Value = True
End Sub

Private Sub optSize_Click(Index As Integer)
    Call DoSample
End Sub
Private Sub chkTest_Click(Index As Integer)
    Call DoSample
End Sub


Private Sub EnumerationExample()

    
    ' THERE ARE TWO WAYS YOU CAN ENUMERATE THE CLASSES
    ' 1. Using the FOR NEXT methods and enumerating that way.
    '    This is shown as Example 2 below. It is END safe
    ' 2. Using the FOR EACH method which is not END safe when uncompiled.
    '    When using this method, include error trapping in the loop
    '    This is shown as Example 1 below
    
    Const ExampleToUse = 1 ' change to 2 for 2nd example
    
    Dim lTotal As Long
    
    
    ' **********************************
    ' EXAMPLE 1 - FOR EACH enumeration '
    ' **********************************
    If ExampleToUse = 1 Then
    
        Dim vList As Variant, vImage As Variant
        
        On Error Resume Next
        For Each vList In lvImageList1.ImageLists
            lTotal = lTotal + vList.Images.Count
        Next
        
        With lvImageList1.ImageLists ' Enumerate image lists & print properites
            Debug.Print "Number of image lists:"; .Count; " -- totaling "; lTotal; "images. Compress list(s) if necessary? "; .Compression
            Debug.Print "FYI: This system is GDI+ or zLIB enabled? "; Not .CompressionNeeded; ". List Compression Recommended for Exports? "; .CompressionNeeded
        End With
        
        For Each vList In lvImageList1.ImageLists ' Enumerate images & print properties
            With vList
                Debug.Print
                ' image list index, key
                Debug.Print vbTab; "List"; .Index; " Key: "; .Key,
                ' image list image count and max images allowed
                Debug.Print "Nr Images:"; .Images.Count; "out of"; .Images.MaxImages
                Debug.Print vbTab; vbTab; "Width:"; .Width; " Height:"; .Height; " DelayLoaded? "; .Images.DelayLoaded
                    
                For Each vImage In vList.Images
                    ' image item's index, key and tag
                    With vImage
                        Debug.Print vbTab; vbTab; "Image"; .Index, _
                        "Key: "; IIf(.Key = vbNullString, "[nothing]", .Key); _
                        "  Tag: "; IIf(.Tag = vbNullString, "[nothing]", .Tag)
                    End With
                Next
            End With
        Next
        
    
    ElseIf ExampleToUse = 2 Then
    
    ' **********************************
    ' EXAMPLE 2 - FOR NEXT enumeration '
    ' **********************************
    
        Dim Index As Long, iImage As Long
    
        With lvImageList1.ImageLists    ' Enumerate image lists and print properties
            For Index = 1 To .Count
                lTotal = lTotal + .Item(Index).Images.Count
            Next
            ' total list count, images, and compression options
            Debug.Print "Number of image lists:"; .Count; " -- totaling "; lTotal; "images. Compress list(s) if necessary? "; .Compression
            Debug.Print "FYI: This system is GDI+ or zLIB enabled? "; Not .CompressionNeeded; ". List Compression Recommended for Exports? "; .CompressionNeeded
            
            For Index = 1 To .Count ' Enumerate images and print properties
                Debug.Print
                ' image list index, key
                Debug.Print vbTab; "List"; Index; " Key: "; .Item(Index).Key,
                
                With .Item(Index)
                    ' image list image count and max images allowed
                    Debug.Print "Nr Images:"; .Images.Count; "out of"; .Images.MaxImages
                    Debug.Print vbTab; vbTab; "Width:"; .Width; " Height:"; .Height; " DelayLoaded? "; .Images.DelayLoaded
                    
                    For iImage = 1 To .Images.Count    ' image count on imagelist
                    
                        ' image item's index, key and tag
                        With .Images(iImage)
                            Debug.Print vbTab; vbTab; "Image"; iImage, _
                            "Key: "; IIf(.Key = vbNullString, "[nothing]", .Key); _
                            "  Tag: "; IIf(.Tag = vbNullString, "[nothing]", .Tag)
                        End With
                    Next
                    
                End With
                
            Next
            
        End With
    
    End If
    
    ' If you want to find an image by Key but don't know what image list it resides on or even if it exists ....
    
    Debug.Print: Debug.Print
    Dim sKey As String
    
    sKey = "DTop" ' not case sensitive
    Debug.Print "EXAMPLE 1:  Looking for key ["; sKey; "] that exists. It actually exists in 2 image lists as you will see"
    
    For Index = 1 To lvImageList1.ImageLists.Count
        With lvImageList1(Index)
            iImage = .Images.IsKeyAssigned(sKey)
            If iImage > 0 Then
                Debug.Print vbTab; sKey; " found on Image List #"; Index; "("; .Key; ") Image Index is: "; iImage
            End If
        End With
    Next
    
    Debug.Print
    sKey = "No Can Do" ' not case sensitive
    Debug.Print "EXAMPLE 2:  Looking for key ["; sKey; "] that does NOT exist"
    
    iImage = 0
    For Index = 1 To lvImageList1.ImageLists.Count
        With lvImageList1(Index)
            If .Images.IsKeyAssigned(sKey) > 0 Then
                iImage = Index
                Debug.Print vbTab; "["; sKey; "] found on Image List #"; Index; "("; .Key; ")"
                Exit For
            End If
        End With
    Next
    If iImage = 0 Then Debug.Print vbTab; "["; sKey; "] not found in any of the"; Index - 1; "image lists."

End Sub
