VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "StandAlone Images - Executable Creator"
   ClientHeight    =   3885
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   259
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   434
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   270
      Left            =   5415
      TabIndex        =   7
      Top             =   105
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   5475
      TabIndex        =   5
      Text            =   "Noname"
      ToolTipText     =   "If no Path supply the Actual Path will to be used"
      Top             =   2580
      Width           =   945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Try"
      Height          =   480
      Left            =   3735
      TabIndex        =   4
      Top             =   3150
      Width           =   1155
   End
   Begin VB.CommandButton CREATE_STANDALONE 
      Caption         =   "Create Executable"
      Height          =   540
      Left            =   3375
      TabIndex        =   3
      Top             =   2430
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   5235
      ScaleHeight     =   555
      ScaleWidth      =   855
      TabIndex        =   2
      Top             =   5295
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CheckBox Check1 
      Caption         =   "View Selection"
      Height          =   210
      Left            =   2625
      TabIndex        =   1
      Top             =   105
      Width           =   1515
   End
   Begin VB.ListBox List1 
      Height          =   1860
      Left            =   2625
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   375
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "File Name"
      Height          =   195
      Left            =   5595
      TabIndex        =   6
      Top             =   2355
      Width           =   840
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1800
      Left            =   135
      Stretch         =   -1  'True
      Top             =   450
      Width           =   2400
   End
   Begin VB.Menu Browse_folder 
      Caption         =   "Browse folder"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Index_folder As Integer
Private Name_to_Save As String
Private path_target As String
Private Ant_listindex As Integer
Private Ant_itemdata As Integer
Private zz As String
Private xxx As String
Private Ant_selected As String
Private Dont_process As Boolean
Private Ant_item As String
Private Pic_nr As Integer
Private Actual_picture As Integer
Private Folder() As String

Private Const BIF_RETURNONLYFSDIRS As Long = &H1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
                          (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
                          (lpBrowseInfo As BROWSEINFO) As Long
Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
                          
Private Sub Browse_folder_Click()

  Dim bi As BROWSEINFO
  Dim IDL As ITEMIDLIST
  Dim pidl As Long
  Dim r As Long
  Dim pos As Integer
  Dim spath As String
  
    If Check1.Value = 1 Then
        Check1.Value = 0
    End If
  
    bi.hOwner = Me.hWnd
       
    bi.pidlRoot = 0&
       
    bi.lpszTitle = "Select one Folder with Graphic files"
       
    bi.ulFlags = BIF_RETURNONLYFSDIRS
       
    pidl& = SHBrowseForFolder(bi)
       
    spath$ = Space$(512)
       
    r = SHGetPathFromIDList(ByVal pidl&, ByVal spath$)

    If r Then
        pos = InStr(spath$, Chr$(0))
        path_target = Left$(spath$, pos - 1)
        If Right$(path_target, 1) = "\" Then
            path_target = Left$(path_target, Len(path_target) - 1)
        End If
      Else
        Exit Sub
    End If
    
    ReDim Preserve Folder(Index_folder)
    Folder(Index_folder) = path_target
    Do_list path_target, Index_folder
    Index_folder = Index_folder + 1

End Sub

Private Sub Check1_Click()

  Dim i As Integer
  Static Condition() As Boolean


    If List1.ListCount = 0 And Check1.Value = 1 Then
        Check1.Value = 0
        Exit Sub
    End If
    If Dont_process Then
        Exit Sub
    End If

    If Check1 Then
        ReDim Condition(List1.ListCount - 1)
        For i = 0 To List1.ListCount - 1
            Condition(i) = List1.Selected(i)
        Next i
again:
        For i = 0 To List1.ListCount - 1
            If Not List1.Selected(i) Then
                List1.RemoveItem i
                GoTo again
            End If
        Next i
      Else
      
        List1.Clear
        For i = 0 To Index_folder - 1
        Do_list Folder(i), i
        Next
        For i = 0 To List1.ListCount - 1
            List1.Selected(i) = Condition(i)
        Next i
    End If

End Sub

Private Sub Command1_Click()
List1.Clear
Erase Folder
Index_folder = 0
End Sub

Private Sub Command2_Click()

  Dim Name As String
  Dim Path As String
  Dim i As Integer

    For i = Len(Name_to_Save) To 1 Step -1
        If Mid$(Name_to_Save, i, 1) = "\" Then
            Name = Mid$(Name_to_Save, i + 1)
            Path = Left$(Name_to_Save, i)
            Exit For
        End If
    Next i

    ShellExecute hWnd, "open", Name, "", Path, 0

End Sub

Private Sub CREATE_STANDALONE_Click()

  Dim free1 As Integer
  Dim free2 As Integer
  Dim Bgr_data_file() As Byte
  Dim i As Integer
  Dim len_file As Long
  Dim Executable() As Byte
  Dim r As Integer
  On Error GoTo erro
  
    If Text1.Text = "" Then
        r = MsgBox("Please enter the Name of the Execubable file", vbCritical)
        Exit Sub
    End If
  
    If List1.SelCount = 0 Then
        r = MsgBox("Please select some file", vbCritical)
        Exit Sub
    End If
  
    If InStr(Text1, ":\") Then
        Name_to_Save = Text1
        If InStr(UCase$(Name_to_Save), ".EXE") = 0 Then
            Name_to_Save = Name_to_Save + ".exe"
        End If
      Else
        Name_to_Save = App.Path & "\" & Text1 & ".exe"
    End If
  
    Executable = LoadResData(101, "CUSTOM")
  
    free1 = FreeFile
    If Dir$(Name_to_Save) <> "" Then
        r = MsgBox(Name_to_Save & " exist. Continue?", vbYesNo)
        If r = 7 Then
            Exit Sub
        End If
        Kill Name_to_Save
    End If
    
    Open Name_to_Save For Binary As #free1
    
    Put #free1, 1, Executable
    
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then
            free2 = FreeFile
        
            Open Folder(List1.ItemData(i)) & "\" & List1.List(i) For Binary As #free2
            ReDim Bgr_data_file(LOF(free2)) As Byte
            Get #free2, 1, Bgr_data_file
            len_file = LOF(free2)
            Put #free1, , "==========/"
            Put #free1, , len_file
            Put #free1, , Bgr_data_file
            Close free2
        End If
    Next i
          
    Close #free1
    
    For i = Len(Name_to_Save) To 1 Step -1
        If Mid(Name_to_Save, i, 1) = "\" Then
            Exit For
        End If
    Next i
    
    r = MsgBox("The File " & UCase(Mid(Name_to_Save, i + 1)) & " was created with sucess" & vbCrLf & "Into folder " & UCase(Left(Name_to_Save, i)))
    
exit_sub:
    Exit Sub
erro:
    r = MsgBox(Error)
    Resume exit_sub
    
End Sub

Private Sub List1_Click()
On Error Resume Next

  Dim X As Long

    If List1.Text = "" Then
        Exit Sub
    End If
    
    Picture1.Picture = LoadPicture(Folder(List1.ItemData(List1.ListIndex)) & "\" & List1.Text)
    X = Picture1.Height / Picture1.Width
    Image1.Height = Image1.Width * X
    Image1.Picture = Picture1.Picture

End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If List1.ListCount = 0 Then Exit Sub
    List1.MousePointer = 7
    Ant_listindex = List1.ListIndex
    Ant_item = List1.List(List1.ListIndex)
    Ant_selected = List1.Selected(List1.ListIndex)
    Ant_itemdata = List1.ItemData(List1.ListIndex)
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim Last_Item As String
  Dim Last_Index As Integer
  Dim Ant_value As Integer
  Dim ant_data As Integer
  
    If List1.ListIndex <> Ant_listindex And Button = 1 Then
        Last_Index = List1.ListIndex
        Ant_value = List1.Selected(Last_Index)
        Last_Item = List1.List(Last_Index)
        ant_data = List1.ItemData(Last_Index)
        
        List1.List(Last_Index) = Ant_item
        List1.Selected(Last_Index) = Ant_selected
        List1.ItemData(Last_Index) = Ant_itemdata
        
        List1.List(Ant_listindex) = Last_Item
        List1.Selected(Ant_listindex) = Ant_value
        List1.ItemData(Ant_listindex) = ant_data
        
        Ant_listindex = Last_Index
        Ant_item = List1.List(Last_Index)
        Ant_selected = List1.Selected(Last_Index)
    
    End If

End Sub

Private Sub Do_list(path_target, index As Integer)

  Dim X As String

       
    Pic_nr = 0
    X = Dir$(path_target & "\*.jpg")
    Do While X <> ""
        List1.AddItem X
        List1.ItemData(List1.NewIndex) = index
        DoEvents
        X = Dir
    Loop
           
    X = Dir$(path_target & "\*.bmp")
    Do While X <> ""
        List1.AddItem X
        List1.ItemData(List1.NewIndex) = index
        DoEvents
        X = Dir
    Loop
    
    X = Dir$(path_target & "\*.gif")
    Do While X <> ""
        List1.AddItem X
        List1.ItemData(List1.NewIndex) = index
        DoEvents
        X = Dir
    Loop
    
    Actual_picture = 0
    
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    List1.MousePointer = 0

End Sub


