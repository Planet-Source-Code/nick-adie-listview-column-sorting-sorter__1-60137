VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List Sort Example"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   12810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvwDetail 
      Height          =   4965
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   8758
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlLVSort 
      Left            =   9660
      Top             =   4020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":01CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":029E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'---------------------------------------------------------------------------------------
' Module    : frmMain
' DateTime  : 21 Apr 2005 15:03
' Author    : Nicholas Adie
' Company   : Brokenmould Limited
' Web       : www.nickadie.com & www.brokenmould.com
' Purpose   : Applications Main Form
'---------------------------------------------------------------------------------------

' Object Definition
' -----------------
Private oLVH                            As CListViewManipulator

 _
 
Private Sub Form_Load()
    
    'Initialise ListView
    '-------------------
    ListViewInitialise
    
    'Load ListView Data
    '------------------
    ListViewLoad

    'Initialise ListView Helper
    '--------------------------
    ' - Initialise Helper Class
    Set oLVH = New CListViewManipulator
    ' - Assign Column Header Icons
    lvwDetail.ColumnHeaderIcons = imlLVSort
    ' - Link ListView To Helper & Sort First Column
    With oLVH
        .Initialise lvwDetail
        .SortColumnContent 1
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)

    'Destroy the Helper
    '------------------
    Set oLVH = Nothing
    
End Sub
Private Sub lvwDetail_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'Sort Any of the Column Headers Clicked
    '--------------------------------------
    oLVH.SortColumnContent ColumnHeader.Index

End Sub
Private Sub ListViewInitialise()

    'Initialise Column Headers
    '-------------------------
    With lvwDetail
    
        ' ~ Standard Settings
        .ColumnHeaders.Add , , "ID", 500
        .ColumnHeaders.Add , , "Date", 1750
        .ColumnHeaders.Add , , "Population (M)", 1500
        .ColumnHeaders.Add , , "Income", 1500
        .ColumnHeaders.Add , , "Code", 1000
        .ColumnHeaders.Add , , "Description", 5000

        ' ~ Set Column Haeder Tags (Required for Non Text Sort)
        .ColumnHeaders(1).Tag = "NUMERIC"
        .ColumnHeaders(2).Tag = "DATE"
        .ColumnHeaders(3).Tag = "NUMERIC"
        .ColumnHeaders(4).Tag = "NUMERIC"
        
        ' ~ Set Column Alignment
        .ColumnHeaders(1).Alignment = lvwColumnLeft
        .ColumnHeaders(2).Alignment = lvwColumnCenter
        .ColumnHeaders(3).Alignment = lvwColumnRight
        .ColumnHeaders(4).Alignment = lvwColumnRight
        .ColumnHeaders(5).Alignment = lvwColumnCenter
        .ColumnHeaders(6).Alignment = lvwColumnLeft
        
        
    End With
    
    'Initialise ListView
    '-------------------
    With lvwDetail
        
        ' ~ Misc Settings
        .MultiSelect = False
        .LabelEdit = lvwManual
        .LabelWrap = False
        .FullRowSelect = True
        .View = lvwReport
        
    End With
    
End Sub
Private Sub ListViewLoad()

    'Definitions
    '-----------
    Dim lngFile                         As Long
    Dim strFile                         As String
    Dim varInput(1)                     As Variant
    Dim Item                            As MSComctlLib.ListItem

    'Initialisation
    '--------------
    ' - Set Randomizer
    Randomize CLng(Format$(Now(), "mmddhhnnss"))
    ' - Get File reference
    lngFile = FreeFile
    ' - Get Data File
    strFile = App.Path & "\data\data.txt"

    Open strFile For Input As #lngFile

    Do
    
        ' - Get Data
        Input #lngFile, varInput(0), varInput(1)
        ' _ Check for end of File
        If EOF(lngFile) Then Exit Do
        
        ' - Create List Item
        Set Item = lvwDetail.ListItems.Add(, , CLng(Rnd * 100))
        Item.SubItems(1) = Format$(CLng(Rnd * 27) + 1 & "/" & CLng(Rnd * 11) + 1 & "/" & 2000 + CLng(Rnd * 5), "dd mmm yyyy")
        Item.SubItems(2) = CLng(Rnd * 1000)
        Item.SubItems(3) = Format(CCur(Rnd * 100), "#,##0.00")
        Item.SubItems(4) = CStr(varInput(0))
        Item.SubItems(5) = CStr(varInput(1))
        
    Loop
        
    Close #lngFile

End Sub
