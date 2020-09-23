VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.UserControl UserControl1 
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7635
   ScaleHeight     =   5895
   ScaleWidth      =   7635
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3885
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6555
      ExtentX         =   11562
      ExtentY         =   6853
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   1950
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Text            =   "UserControl1.ctx":0000
      Top             =   2580
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CommandButton Command1 
      Height          =   315
      Left            =   6450
      Picture         =   "UserControl1.ctx":085F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   375
   End
   Begin VB.TextBox txtAddress 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   330
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   0
      Width           =   6105
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   900
      Pattern         =   "*.jpg;*.gif;*.bmp"
      TabIndex        =   0
      Top             =   4350
      Visible         =   0   'False
      Width           =   2835
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSelect 
         Caption         =   "Select"
      End
      Begin VB.Menu mnuSetProp 
         Caption         =   "Set Properties"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom"
      End
   End
   Begin VB.Menu mnuGeneral 
      Caption         =   "General"
      Begin VB.Menu mnuChangeLocation 
         Caption         =   "Change Location..."
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetTempPath _
    Lib "kernel32" Alias "GetTempPathA" _
   (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long

Private Declare Function GetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
   (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String) As Long

Private Const MAX_PATH As Long = 260

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const BIF_DONTGOBELOWDOMAIN = 2

Private Type BrowseInfo
   hwndOwner      As Long
   pIDLRoot       As Long
   pszDisplayName As Long
   lpszTitle      As Long
   ulFlags        As Long
   lpfnCallback   As Long
   lParam         As Long
   iImage         As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long

Dim mDoc As HTMLDocument
Dim imgHeightWidth As Long
Dim imgHeightWidthCont As Long
Dim diff2Images As Long
Dim WithEvents doc As HTMLDocument
Attribute doc.VB_VarHelpID = -1
Const RIGHT_BUTTON = 2
Public Event dblclick(path As String)
Public Event click(path As String)
Public Event mouseover(path As String)

Private FormLoaded As Boolean
Dim alt As String
Dim TempFileName As String
'

Private Function OpenDirectoryTV(ohwnd As Long, Optional odtvTitle As String) As String
   Dim lpIDList As Long
   Dim sBuffer As String
   Dim szTitle As String
   Dim tBrowseInfo As BrowseInfo
   szTitle = odtvTitle
   With tBrowseInfo
      .hwndOwner = ohwnd
      .lpszTitle = lstrcat(szTitle, "")
      .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
   End With
   lpIDList = SHBrowseForFolder(tBrowseInfo)
   If (lpIDList) Then
      sBuffer = Space(MAX_PATH)
      SHGetPathFromIDList lpIDList, sBuffer
      sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
      OpenDirectoryTV = sBuffer
   End If
End Function

Private Sub Command2_Click()
    WebBrowser1.Width = WebBrowser1.Width - 60
End Sub

Private Sub Command1_Click()
    Dim strPath As String
    strPath = OpenDirectoryTV(UserControl.hWnd, "Select image folder")
    If Len(Trim(strPath)) = 0 Then Exit Sub
    txtAddress.Text = strPath
    setPath
End Sub

Private Sub doc_ondataavailable()
    MsgBox "downloaded"
End Sub

Private Function doc_ondblclick() As Boolean
    Dim eventObj As IHTMLEventObj
    Dim srcName As String
    
    Set eventObj = doc.parentWindow.event
    alt = ""
    srcName = eventObj.srcElement.tagName
    
    If srcName = "IMG" Or srcName = "SPAN" Or srcName = "INPUT" Or srcName = "DIV" Then
        alt = eventObj.srcElement.getAttribute("ID")
        RaiseEvent dblclick(alt)
    End If
    
End Function

Private Function doc_onclick() As Boolean
    Dim eventObj As IHTMLEventObj
    Dim srcName As String
    
    Set eventObj = doc.parentWindow.event
    alt = ""
    srcName = eventObj.srcElement.tagName
    
    If srcName = "IMG" Or srcName = "SPAN" Or srcName = "INPUT" Or srcName = "DIV" Then
        alt = eventObj.srcElement.getAttribute("ID")
        RaiseEvent click(alt)
    End If
    
End Function

Private Sub doc_onmousedown()
    Dim eventObj As IHTMLEventObj
    Dim srcName As String
    
    Set eventObj = doc.parentWindow.event
    alt = ""
    srcName = eventObj.srcElement.tagName
    
    If eventObj.button = RIGHT_BUTTON Then
        If srcName = "IMG" Or srcName = "SPAN" Or srcName = "INPUT" Or srcName = "DIV" Then
            alt = eventObj.srcElement.getAttribute("ID")
            If Len(Trim(alt)) > 0 Then
                PopupMenu mnuFile
            Else
                PopupMenu mnuGeneral
            End If
        Else
            On Error Resume Next
            PopupMenu mnuGeneral
        End If
    End If
    
'    If eventObj.button = 1 Then
'        If srcName = "IMG" Or srcName = "SPAN" Or srcName = "INPUT" Or srcName = "DIV" Then
'            alt = eventObj.srcElement.getAttribute("ID")
'            RaiseEvent ImageSelect(alt)
'        End If
'    End If
End Sub

Private Sub doc_onmouseover()
    Dim eventObj As IHTMLEventObj
    Dim srcName As String
    
    Set eventObj = doc.parentWindow.event
    alt = ""
    srcName = eventObj.srcElement.tagName
    
    If srcName = "IMG" Or srcName = "SPAN" Or srcName = "INPUT" Or srcName = "DIV" Then
        alt = eventObj.srcElement.getAttribute("ID")
        RaiseEvent mouseover(alt)
    End If
End Sub

Private Sub mnuSelect_Click()
    RaiseEvent dblclick(alt)
End Sub
Private Sub mnuZoom_Click()
    frmZoom.OpenImage (alt)
    frmZoom.Show vbModal
End Sub
Private Sub mnuRefresh_Click()
    WebBrowser1.Refresh
    WebBrowser1.Refresh2
End Sub
Private Sub mnuChangeLocation_Click()
    Command1_Click
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        setPath
    End If
End Sub

Private Sub usercontrol_initialize()
    TempFileName = GetSystemTempPath & "34rxdt92[1].html"
    imgHeightWidth = 100
    imgHeightWidthCont = imgHeightWidth + 10
    diff2Images = 20
    
    txtAddress.Text = App.path
    Set mDoc = New HTMLDocument
    WebBrowser1.Navigate2 "about:blank"
    'setPath
End Sub
Public Sub SetLocation(mFolder As String)
    txtAddress.Text = mFolder
    setPath
End Sub
Private Sub setPath()
    Dim ImageFolderPath As String
    Dim mDocContents As String
    ImageFolderPath = txtAddress.Text
    CreateFile (ImageFolderPath)
    'Picture1.Enabled = False
    If PathExists(ImageFolderPath) Then
        WebBrowser1.Navigate2 "file:///" & TempFileName
    End If
End Sub

Private Function CreateFile(ImageFolderPath As String) As String
    Dim retValue As String
    Dim tempFile As String
    Dim i As Integer
    Dim ID As String
    
    If Not PathExists(ImageFolderPath) Then
        MsgBox "Path does not exists", vbCritical, "Path Error"
        Exit Function
    End If
    
    File1.path = ImageFolderPath
    
    UserControl.MousePointer = vbHourglass
    UserControl.Enabled = False
    retValue = ""
    For i = 0 To File1.ListCount - 1
        ID = File1.path & "\" & File1.List(i)
        
        retValue = retValue & _
                    "<span readonly class='borderspan' style='width:" & imgHeightWidthCont + diff2Images & ";height:" & imgHeightWidthCont + diff2Images + 20 & "'>" & _
                    "<span id='" & ID & "' readonly onmouseout='mOut(this)' onmouseover='mOver(this)' class='border' style='cursor:hand;width:" & (imgHeightWidthCont) & ";height:" & (imgHeightWidthCont) & "'>" & _
                    "<img src='" & ID & "' onload='setSize(this)' id='" & ID & "' style='cursor:hand;'>" & _
                    "</span>" & _
                    "<input id='" & ID & "' type='text' class='noborder' style='width:" & imgHeightWidthCont & ";cursor:hand;text-align:center;color:#ffffff;background-color:#666666' readonly value='" & File1.List(i) & "' align='right'>" & _
                    "</span>" & vbCrLf
    
'        retValue = retValue & "<span readonly id='" & File1.path & "\" & File1.List(i) & "' onmouseout='mOut(this)' onmouseover='mOver(this)' align='center' valign='middle' class='border' style='cursor:hand;width:" & (imgHeightWidth + 20) & ";height:" & (imgHeightWidth + 20) & "'>" & _
                    "<center id='" & File1.path & "\" & File1.List(i) & "'>" & _
                    "<img src='" & File1.path & "\" & File1.List(i) & "' onload='setSize(this)' id='" & File1.path & "\" & File1.List(i) & "' style='cursor:hand;' align='center' valign='absmiddle'><br><input id='" & File1.path & "\" & File1.List(i) & "' type='text' class='noborder' style='cursor:hand;text-align:center;color:#ffffff;background-color:#666666' size='" & imgHeightWidth \ 6 & "' readonly value='" & File1.List(i) & "' align='right'>" & _
                    "</center></span>&nbsp;&nbsp;" & vbCrLf
                    DoEvents

    Next
    
    retValue = Replace(Text2.Text, "##IMAGE_DETAILS##", retValue)
    
    'Attributes
    '##IMAGE_BORDER_COLOR##
    '##HIGHLIGHT_BACKGROUND_COLOR##
    '##HIGHLIGHT_IMAGE_BORDER_COLOR##
    '##BODY_BGCOLOR##
    
    'bodycolor
    'retValue = Replace(retValue, "##BODY_BGCOLOR##", "#DDDDDD")
    retValue = Replace(retValue, "##BODY_BGCOLOR##", "#FFFFFF")
    
    'Image border color
    retValue = Replace(retValue, "##IMAGE_BORDER_COLOR##", "#AAAAAA")
    
    'Image Highlight color
    retValue = Replace(retValue, "##HIGHLIGHT_BACKGROUND_COLOR##", "#999999")
    
    'HIGHLIGHT_IMAGE_BORDER_COLOR
    retValue = Replace(retValue, "##HIGHLIGHT_IMAGE_BORDER_COLOR##", "#333333")
    
    'SIZE_HEIGHT
    retValue = Replace(retValue, "##SIZE_HEIGHT_WIDTH##", imgHeightWidth)
    
    'SIZE_HEIGHT
    retValue = Replace(retValue, "##SIZE_HEIGHT_WIDTH_CONT##", imgHeightWidthCont)
    
    'difference in width of 2 images
    retValue = Replace(retValue, "##DIFF_2_IMAGES##", diff2Images)
    
    'SIZE_HEIGHT
    retValue = Replace(retValue, "##OUTER_SPAN_SIZE##", imgHeightWidthCont + diff2Images)
    
    
    Open TempFileName For Output As #1
    Print #1, retValue
    Close #1
    
End Function

Private Sub usercontrol_Resize()
    On Error Resume Next
    txtAddress.Move 0, 0, ScaleWidth - Command1.Width - 60, txtAddress.Height
    Command1.Move txtAddress.Left + txtAddress.Width + 30, 0
    WebBrowser1.Move 0, txtAddress.Top + txtAddress.Height + 15, UserControl.Width, UserControl.Height - WebBrowser1.Top
    'WebBrowser1.Move -30, -30, Picture1.Width + 45, Picture1.Height + 45
End Sub


Public Property Get ImageHeightWidth() As Variant
    ImageHeightWidth = imgHeightWidth
End Property

Public Property Let ImageHeightWidth(ByVal vNewValue As Variant)
    If vNewValue < 50 Then vNewValue = 50
    imgHeightWidth = vNewValue
End Property

Private Sub UserControl_Terminate()
    On Error Resume Next
    Kill TempFileName
End Sub

Private Sub WebBrowser1_CommandStateChange(ByVal Command As Long, ByVal Enable As Boolean)
    On Error Resume Next
    'Clear the selection
    If FormLoaded Then
        WebBrowser1.ExecWB OLECMDID_CLEARSELECTION, OLECMDEXECOPT_DONTPROMPTUSER
    End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    FormLoaded = True
    Set doc = WebBrowser1.Document
    UserControl.MousePointer = vbNormal
    UserControl.Enabled = True
End Sub


Private Function PathExists(mPath As String) As Boolean
    On Error GoTo errh
    File1.path = mPath
    PathExists = True
    Exit Function
errh:
    PathExists = False
End Function


Public Function GetSystemTempPath() As String
   
   Dim result As Long
   Dim buff As String
   
   buff = Space$(MAX_PATH)
   result = GetTempPath(MAX_PATH, buff)
   GetSystemTempPath = Left$(buff, result)
End Function



