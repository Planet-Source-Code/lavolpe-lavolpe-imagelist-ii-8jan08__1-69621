VERSION 5.00
Begin VB.UserControl lvImageList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   600
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "lvImageList.ctx":0000
   PropertyPages   =   "lvImageList.ctx":11EE
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   40
   ToolboxBitmap   =   "lvImageList.ctx":1201
End
Attribute VB_Name = "lvImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' UserControl is only an IDE interface for the Image List classes

' NOTE: TO CONVERT USERCONTROL TO CLASS ONLY, RUN-TIME ONLY IMAGELIST
' - Simply remove the usercontrol and property page from the project
' - Add the classes and and bas module to your project
' - Call ImageLists.ManageGDItoken during Form_Load and pass it the form's hWnd
' - can safely remove the following functions used only by the property page
'       cListImages.xcom_Clone, .xcom_CloneData, .xcom_IsCloned methods
'                  .xcom_AddImage_FromDIB, .xcom_ImagesDIB
' - The remaining xcom_.... properties and methods will become available to you.
'   Do not call them from within your project. They are for internal
'   class to class communications only
' - VB may change some class instancing properties on you; that is ok

Private m_Lists As cImageLists

Public Property Get ImageLists() As cImageLists
Attribute ImageLists.VB_UserMemId = 0
Attribute ImageLists.VB_MemberFlags = "200"
    ' returns the imagelists collection
    Set ImageLists = m_Lists
    ' IMPORTANT: This property must be set to Procedure ID: Default and also
    ' checked as User Interface Default. See menu: Tools|Procedure Attributes, Advanced
    ' Moving this routine may reset those properites: double check them
End Property

Friend Property Let UpdateLists(bUpdate As Boolean)
    ' internal use only & called by the property page
    ' to inform usercontrol that data has changed
    PropertyChanged "UpdateLists"
End Property


Private Sub UserControl_Initialize()
    Set m_Lists = New cImageLists
End Sub

Private Sub UserControl_InitProperties()
    m_Lists.ManageGDIToken ContainerHwnd ' provide GDI+ safe environment
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    
    m_Lists.ManageGDIToken ContainerHwnd ' provide GDI+ safe environment
    With PropBag
        Dim bData() As Byte, nrLists As Long

        nrLists = .ReadProperty("Count", vbDefault)    ' how many imagelists are expected?
        If nrLists Then
            On Error Resume Next
            For nrLists = 1 To nrLists
                bData() = .ReadProperty("Img" & nrLists, vbNullString) ' retrieve saved imagelist data
                If Err Then
                    Err.Clear
                Else
                    m_Lists.ImportImageList bData()         ' pass array off
                End If
            Next
        End If
    End With

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
        
    With PropBag
        If m_Lists.Count Then        ' store the various image lists
            Dim Index As Long, bData() As Byte
            Dim nrLists As Long
            On Error Resume Next
            For Index = 1 To m_Lists.Count
                If m_Lists.ExportImageList(Index, bData()) Then
                    If Err Then
                        Err.Clear
                    Else
                        nrLists = nrLists + 1
                        .WriteProperty "Img" & nrLists, bData()
                    End If
                Else
                    ' don't save null lists
                End If
            Next
        End If
        .WriteProperty "Count", nrLists, vbDefault ' save number of imagelists cached
    End With

End Sub
