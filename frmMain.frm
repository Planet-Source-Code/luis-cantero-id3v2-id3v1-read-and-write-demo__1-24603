VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ID3v2 Read/Write Demo"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdWriteID3v2 
      Caption         =   "Write ID3v2"
      Height          =   495
      Left            =   6240
      TabIndex        =   16
      Top             =   4920
      Width           =   1215
   End
   Begin VB.CommandButton cmdWriteID3v1 
      Caption         =   "Write ID3v1"
      Height          =   495
      Left            =   6240
      TabIndex        =   15
      Top             =   4320
      Width           =   1215
   End
   Begin VB.FileListBox FilOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   3120
      OLEDropMode     =   1  'Manual
      Pattern         =   "*.mp3"
      System          =   -1  'True
      TabIndex        =   14
      ToolTipText     =   "Click to Select or Double Click to Open"
      Top             =   480
      Width           =   4365
   End
   Begin VB.DriveListBox DrvOpen 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3000
   End
   Begin VB.DirListBox DirOpen 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   12
      Top             =   480
      Width           =   3000
   End
   Begin VB.TextBox txtMP3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Index           =   4
      Left            =   1275
      MaxLength       =   254
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   7
      Top             =   4110
      Width           =   2325
   End
   Begin VB.TextBox txtMP3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Index           =   3
      Left            =   1275
      MaxLength       =   254
      OLEDragMode     =   1  'Automatic
      OLEDropMode     =   2  'Automatic
      TabIndex        =   6
      Top             =   3795
      Width           =   2325
   End
   Begin VB.TextBox txtMP3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Index           =   2
      Left            =   1275
      Locked          =   -1  'True
      MaxLength       =   254
      TabIndex        =   5
      Top             =   3465
      Width           =   2325
   End
   Begin VB.TextBox txtMP3 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   250
      Index           =   1
      Left            =   1275
      Locked          =   -1  'True
      MaxLength       =   254
      TabIndex        =   4
      Top             =   3150
      Width           =   2325
   End
   Begin VB.CheckBox chkUnkTitle 
      Caption         =   "If not in tag use: ""Untitled"""
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "If the Title is not in then Tag the word ""Untitled"" will be used"
      Top             =   4080
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.CheckBox chkUcase 
      Caption         =   "Convert First Letter To Uppercase"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3630
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3435
      Value           =   1  'Checked
      Width           =   2535
   End
   Begin VB.OptionButton opID3 
      Caption         =   "ID3v1"
      Height          =   300
      Index           =   0
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4440
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.OptionButton opID3 
      Height          =   300
      Index           =   1
      Left            =   5550
      Picture         =   "frmMain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "ID3v2"
      Top             =   4440
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblToDo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed Title:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   10
      Left            =   120
      TabIndex        =   11
      Top             =   4125
      Width           =   1065
   End
   Begin VB.Label lblToDo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Proposed Artist:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   10
      Top             =   3810
      Width           =   1110
   End
   Begin VB.Label lblToDo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Title:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   9
      Top             =   3495
      Width           =   930
   End
   Begin VB.Label lblToDo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Existing Artist:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   8
      Top             =   3180
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function ReadID3v1() As String

  Dim StrGanzerTag As String * 128
  Dim StrSongName As String * 30
  Dim StrArtist As String * 30

    On Error GoTo Problems
    tmpFile = FreeFile
    If FilOpen = "" Then Exit Function
    If Len(DirOpen) = 3 Then tmpDirOpen = Left$(DirOpen, 2) Else tmpDirOpen = DirOpen
    Open tmpDirOpen & "\" & FilOpen For Binary Access Read As tmpFile
    If LOF(tmpFile) > 127 Then
        Seek tmpFile, LOF(tmpFile) - 127
        Get tmpFile, , StrGanzerTag
        Close tmpFile
    End If

    If InStrB(StrGanzerTag, "TAG") = 1 Then
        StrSongName = Mid$(StrGanzerTag, 4, 30)
        StrArtist = Mid$(StrGanzerTag, 34, 30)
        StrAlbum = Mid$(StrGanzerTag, 64, 30)
        StrAlbum = Mid$(StrAlbum, 1, 31 - (InStr(1, StrAlbum, Chr$(0))))
        If StrAlbum = String$(30, 0) Then StrAlbum = ""
        StrTrack = Asc(Mid$(StrGanzerTag, 127, 1))

        Artist = Trim$(StrArtist)
        Title = Trim$(StrSongName)

        txtMP3(1) = Artist
        txtMP3(2) = Title
        txtMP3(3) = Artist
        txtMP3(4) = Title

        If chkUcase = 1 Then
            For Index = 3 To 4
                For i = 1 To Len(txtMP3(Index))
                    If i = 1 Then strTemp = UCase$(Mid$(txtMP3(Index), i, 1))
                    If (Mid$(txtMP3(Index), i, 1) = " " Or Mid$(txtMP3(Index), i, 1) = "(" Or Mid$(txtMP3(Index), i, 1) = "." Or Mid$(txtMP3(Index), i, 1) = "-") And i < Len(txtMP3(Index)) Then
                        strTemp = strTemp & UCase$(Mid$(txtMP3(Index), i + 1, 1))
                      Else
                        strTemp = strTemp & Mid$(txtMP3(Index), i + 1, 1)
                    End If
                Next
                txtMP3(Index) = strTemp
            Next
        End If
        ReadID3v1 = txtMP3(3) & "§ " & txtMP3(4) & "§ " & StrAlbum & "§ " & StrTrack
      Else
        If chkUnkArtist = 1 Then Artist = "Unknown" Else Artist = ""
        If chkUnkTitle = 1 Then Title = "Untitled" Else Title = ""
        txtMP3(1) = "": txtMP3(2) = ""
        txtMP3(3) = Artist: txtMP3(4) = Title
    End If

Exit Function

Problems:
    Close tmpFile
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Function

Private Function ReadID3v2() As String

  Dim StrGanzerTag As String * 10
  Dim bytGanzerTag() As Byte
  Dim strIDdv2Tag As String
  Dim strTagLength As Long

    On Error GoTo Problems
    tmpFile = FreeFile
    
    If FilOpen = "" Then Exit Function
    If Len(DirOpen) = 3 Then tmpDirOpen = Left$(DirOpen, 2) Else tmpDirOpen = DirOpen
    Open tmpDirOpen & "\" & FilOpen For Binary As tmpFile
    Get tmpFile, , StrGanzerTag

    If Left$(StrGanzerTag, 3) = "ID3" Then ' v2 Tag exist
        strTagLength = Asc(Mid$(StrGanzerTag, 7, 1)) * 2 ^ 21 + Asc(Mid$(StrGanzerTag, 8, 1)) * CLng(2 ^ 14) + Asc(Mid$(StrGanzerTag, 9, 1)) * 2 ^ 7 + Asc(Mid$(StrGanzerTag, 10, 1)) * 2 ^ 0
        
        ReDim bytGanzerTag(strTagLength) As Byte
        Seek tmpFile, 1
        Get tmpFile, , bytGanzerTag
        Close tmpFile
        
        strIDdv2Tag = StrConv(bytGanzerTag, vbUnicode)

        tmp1 = InStr(11, strIDdv2Tag, "TPE1" & String$(3, 0))
        tmp2 = Asc(Mid$(strIDdv2Tag, tmp1 + 7, 1)) - 1
        tmp3 = InStr(11, strIDdv2Tag, "TIT2" & String$(3, 0))
        tmp4 = Asc(Mid$(strIDdv2Tag, tmp3 + 7, 1)) - 1
        tmp5 = InStr(11, strIDdv2Tag, "TALB" & String$(3, 0))
        tmp6 = Asc(Mid$(strIDdv2Tag, tmp5 + 7, 1)) - 1
        tmp7 = InStr(11, strIDdv2Tag, "TRCK" & String$(3, 0))
        tmp8 = Asc(Mid$(strIDdv2Tag, tmp7 + 7, 1)) - 1

        If tmp1 = 0 And tmp3 = 0 Then
            If ReadID3v1 = "" Then
                If chkUnkArtist = 1 Then Artist = "Unknown" Else Artist = ""
                If chkUnkTitle = 1 Then Title = "Untitled" Else Title = ""
                txtMP3(1) = "": txtMP3(2) = ""
                txtMP3(3) = Artist: txtMP3(4) = Title
            End If
            Exit Function
        End If
    
        StrArtist = Mid$(strIDdv2Tag, tmp1 + 11, tmp2)
        StrSongName = Mid$(strIDdv2Tag, tmp3 + 11, tmp4)
        If tmp5 > 0 Then StrAlbum = Mid$(strIDdv2Tag, tmp5 + 11, tmp6)
        If tmp7 > 0 Then StrTrack = Mid$(strIDdv2Tag, tmp7 + 11, tmp8)

        Artist = Trim$(StrArtist)
        Title = Trim$(StrSongName)
        
        txtMP3(1) = Artist
        txtMP3(2) = Title
        txtMP3(3) = Artist
        txtMP3(4) = Title

        If chkUcase = 1 Then
            For Index = 3 To 4
                For i = 1 To Len(txtMP3(Index))
                    If i = 1 Then strTemp = UCase$(Mid$(txtMP3(Index), i, 1))
                    If (Mid$(txtMP3(Index), i, 1) = " " Or Mid$(txtMP3(Index), i, 1) = "(" Or Mid$(txtMP3(Index), i, 1) = "." Or Mid$(txtMP3(Index), i, 1) = "-") And i < Len(txtMP3(Index)) Then
                        strTemp = strTemp & UCase$(Mid$(txtMP3(Index), i + 1, 1))
                      Else
                        strTemp = strTemp & Mid$(txtMP3(Index), i + 1, 1)
                    End If
                Next
                txtMP3(Index) = strTemp
            Next
        End If
        ReadID3v2 = txtMP3(3) & "§ " & txtMP3(4) & "§ " & StrAlbum & "§ " & StrTrack
      Else
        If ReadID3v1 = "" Then
            If chkUnkArtist = 1 Then Artist = "Unknown" Else Artist = ""
            If chkUnkTitle = 1 Then Title = "Untitled" Else Title = ""
            txtMP3(1) = "": txtMP3(2) = ""
            txtMP3(3) = Artist: txtMP3(4) = Title
        End If
    End If
    Close tmpFile

Exit Function

Problems:
    Close tmpFile
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number
    lblStatus = "Error number " & Err.Number & " - " & Err.Description

End Function

Private Sub WriteID3v1()

  Dim StrGanzerTag As String * 128

    On Error GoTo Problems
    txtMP3(3) = Left$(txtMP3(3), 30) 'ID3v1 only suports 30 chars
    txtMP3(4) = Left$(txtMP3(4), 30)
    
    tmpFile = FreeFile
    If Len(DirOpen) = 3 Then tmpDirOpen = Left$(DirOpen, 2) Else tmpDirOpen = DirOpen
    Open tmpDirOpen & "\" & FilOpen For Binary As tmpFile
    If LOF(tmpFile) > 127 Then
        Seek tmpFile, LOF(tmpFile) - 127
        Get tmpFile, , StrGanzerTag
    End If

    strNewTag = "TAG" & txtMP3(4) & Space$(30 - Len(txtMP3(4))) & txtMP3(3) & Space$(30 - Len(txtMP3(3))) & Right$(StrGanzerTag, 65)
    If InStr(1, Caption, "Unregistered") > 0 Then strNewTag = Left$(strNewTag, Len(strNewTag) - 31) & "http://LCenterprises.net      " & Chr$(255)

    If InStrB(StrGanzerTag, "TAG") = 1 Then ' Tag exist
        Seek tmpFile, LOF(tmpFile) - 127
        If StrGanzerTag <> strNewTag Then Put tmpFile, , strNewTag
      Else ' New ID3v1 Tag
        strNewTag = "TAG" & txtMP3(4) & Space$(30 - Len(txtMP3(4))) & txtMP3(3) & Space$(30 - Len(txtMP3(3))) & Space$(65)
        If InStr(1, Caption, "Unregistered") > 0 Then strNewTag = Left$(strNewTag, Len(strNewTag) - 31) & "http://LCenterprises.net      " & Chr$(255)
        Put tmpFile, , strNewTag
    End If
    Close tmpFile

Exit Sub

Problems:
    Close tmpFile
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Sub

Private Sub WriteID3v2(Optional bolRemoveID3v1 As Boolean)

  Dim StrGanzerTag As String * 10
  Dim strID3v1 As String * 128
  Dim strID3v1Fields As String * 65
  Dim bytGanzerTag() As Byte
  Dim strNewTag As String
  Dim strTagLength As Long

    On Error GoTo Problems
    tmpFile = FreeFile
    If Len(DirOpen) = 3 Then tmpDirOpen = Left$(DirOpen, 2) Else tmpDirOpen = DirOpen
    Open tmpDirOpen & "\" & FilOpen For Binary As tmpFile
    Get tmpFile, , StrGanzerTag 'Read ID3v2 Header
    
    Seek tmpFile, LOF(tmpFile) - 127
    Get tmpFile, , strID3v1 ' Get ID3v1
    If Left$(strID3v1, 3) = "TAG" Then
        strID3v1Fields = Mid$(strID3v1, 64)
    End If

    If Left$(StrGanzerTag, 3) = "ID3" Then ' v2 Tag exist
        strTagLength = Asc(Mid$(StrGanzerTag, 7, 1)) * 2 ^ 21 + Asc(Mid$(StrGanzerTag, 8, 1)) * CLng(2 ^ 14) + Asc(Mid$(StrGanzerTag, 9, 1)) * 2 ^ 7 + Asc(Mid$(StrGanzerTag, 10, 1)) * 2 ^ 0
        
        ReDim bytGanzerTag(strTagLength - 1) As Byte
        Seek tmpFile, 1
        Get tmpFile, , bytGanzerTag
        
        strNewTag = StrConv(bytGanzerTag, vbUnicode)
        strCurrentTag = strNewTag

        tmp1 = InStr(11, strNewTag, "TPE1" & String$(3, 0)) 'Get Artist's position
        
        If tmp1 > 0 Then 'Artist exists in Tag
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(txtMP3(3)) + 1) & String$(3, 0) & txtMP3(3) & Mid$(strNewTag, tmp1 + 11 + tmp2)
          Else 'Artist does not exist in Tag
            tmp1 = InStr(11, strNewTag, String$(Len(txtMP3(3)) + 11, 0)) ' Look for space
            strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "TPE1" & String$(3, 0) & Chr$(Len(txtMP3(3)) + 1) & String$(3, 0) & txtMP3(3) & Mid$(strNewTag, Len(txtMP3(3)) + 11)
        End If

        tmp1 = InStr(11, strNewTag, "TIT2" & String$(3, 0)) 'Get Title's position
                
        If tmp1 > 0 Then 'Title exists in Tag
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(txtMP3(4)) + 1) & String$(3, 0) & txtMP3(4) & Mid$(strNewTag, tmp1 + 11 + tmp2)
          Else 'Title does not exist in Tag
            tmp1 = InStr(11, strNewTag, String$(Len(txtMP3(4)) + 11, 0)) ' Look for space
            strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "TIT2" & String$(3, 0) & Chr$(Len(txtMP3(4)) + 1) & String$(3, 0) & txtMP3(4) & Mid$(strNewTag, tmp1 + Len(txtMP3(4)) + 11)
        End If

        tmp1 = InStr(11, strNewTag, "COMM" & String$(3, 0)) 'Get Comment's position
        tmpComments = Trim$(Mid$(strID3v1Fields, 35, 30))
                
        If tmp1 > 0 Then 'Comment exists in Tag
            
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            
            If InStr(1, Caption, "Unregistered") > 0 Then
                strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(29) & String$(3, 0) & "eng" & Chr$(0) & "http://LCenterprises.net" & Mid$(strNewTag, tmp1 + 11 + tmp2)
              Else
                strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(tmpComments) + 5) & String$(3, 0) & "eng" & Chr$(0) & tmpComments & Mid$(strNewTag, tmp1 + 11 + tmp2)
            End If

          Else 'Comment does not exist in Tag
            If InStr(1, Caption, "Unregistered") > 0 Then
                tmp1 = InStr(11, strNewTag, String$(28 + 11, 0)) ' Look for space
                strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "COMM" & String$(3, 0) & Chr$(29) & String$(3, 0) & "eng" & Chr$(0) & "http://LCenterprises.net" & Mid$(strNewTag, tmp1 + 28 + 11)
              Else
                tmp1 = InStr(11, strNewTag, String$(Len(tmpComments) + 15, 0)) ' Look for space
                strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "COMM" & String$(3, 0) & Chr$(Len(tmpComments) + 5) & String$(3, 0) & "eng" & Chr$(0) & tmpComments & Mid$(strNewTag, tmp1 + Len(tmpComments) + 15)
            End If
        End If
        
        tmp1 = InStr(11, strNewTag, "TALB" & String$(3, 0)) 'Get Album's position
        tmpAlbum = Trim$(Mid$(strID3v1Fields, 31, 4))
                
        If tmp1 > 0 Then 'Album exists in Tag
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(tmpAlbum) + 1) & String$(3, 0) & tmpAlbum & Mid$(strNewTag, tmp1 + 11 + tmp2)
          Else 'Year does not exist in Tag
            tmp1 = InStr(11, strNewTag, String$(Len(tmpYear) + 11, 0)) ' Look for space
            strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "TALB" & String$(3, 0) & Chr$(Len(tmpAlbum) + 1) & String$(3, 0) & tmpAlbum & Mid$(strNewTag, tmp1 + Len(tmpAlbum) + 11)
        End If

        tmp1 = InStr(11, strNewTag, "TYER" & String$(3, 0)) 'Get Year's position
        tmpYear = Trim$(Mid$(strID3v1Fields, 31, 4))
                
        If tmp1 > 0 Then 'Year exists in Tag
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(tmpYear) + 1) & String$(3, 0) & tmpYear & Mid$(strNewTag, tmp1 + 11 + tmp2)
          Else 'Year does not exist in Tag
            tmp1 = InStr(11, strNewTag, String$(Len(tmpYear) + 11, 0)) ' Look for space
            strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "TYER" & String$(3, 0) & Chr$(Len(tmpYear) + 1) & String$(3, 0) & tmpYear & Mid$(strNewTag, tmp1 + Len(tmpYear) + 11)
        End If
        
        tmp1 = InStr(11, strNewTag, "TCON" & String$(3, 0)) 'Get Genre's position
        tmpGenre = Asc(Mid$(strID3v1Fields, 65))

        If tmp1 > 0 Then 'Genre exists in Tag
            tmp2 = Asc(Mid$(strNewTag, tmp1 + 7, 1)) - 1 ' Get size of Tag
            strNewTag = Mid$(strNewTag, 1, tmp1 + 6) & Chr$(Len(tmpGenre) + 3) & String$(3, 0) & "(" & tmpGenre & ")" & Mid$(strNewTag, tmp1 + 11 + tmp2)
          Else 'Genre does not exist in Tag
            tmp1 = InStr(11, strNewTag, String$(Len(tmpGenre) + 13, 0)) ' Look for space
            strNewTag = Mid$(strNewTag, 1, tmp1 - 1) & "TCON" & String$(3, 0) & Chr$(Len(tmpGenre) + 3) & String$(3, 0) & "(" & tmpGenre & ")" & Mid$(strNewTag, tmp1 + Len(tmpGenre) + 13)
        End If
        
        If Len(strNewTag) < strTagLength Then 'New Tag is smaller than Tag length (char. deleted)
            strNewTag = strNewTag & String$(strTagLength - Len(strNewTag), 0)
          Else 'New Tag is bigger than Tag length (char. added)
            strNewTag = Mid$(strNewTag, 1, strTagLength)
        End If
        
        Seek tmpFile, 1

        If bolRemoveID3v1 And Left$(strID3v1, 3) = "TAG" Then 'If ID3v1 should be removed
            Put tmpFile, , strNewTag
            ReDim strFile1(LOF(tmpFile) - 129) As Byte ' Get file without Tag
            Seek tmpFile, 1
            
            lblStatus = "ID3v2 exists, loading file..."
            DoEvents
            
            Get tmpFile, , strFile1
            Close tmpFile
            
            Open tmpDirOpen & "\" & FilOpen For Output As tmpFile ' Erase contents of file
            Close tmpFile
            
            Open tmpDirOpen & "\" & FilOpen For Binary As tmpFile ' Put mp3 data into file
        
            lblStatus = "ID3v2 exists, ID3v1 removed, rewriting file..."
            DoEvents
            Put tmpFile, , strFile1 ' put mp3 data
          Else 'If ID3v1 should be kept or does not exist
            lblStatus = "ID3v2 exists, ID3v1 kept or does not exist, updating file..."
            DoEvents
            
            If strCurrentTag <> strNewTag Then Put tmpFile, , strNewTag
        End If
        
      Else ' New ID3v2 Tag
        strNewTag = "ID3" & Chr$(3) & Chr$(0) & String$(3, 0) & Chr$(2) & Chr$(1) & "TPE1" & String$(3, 0) & Chr$(Len(txtMP3(3)) + 1) & String$(3, 0) & txtMP3(3) & "TIT2" & String$(3, 0) & Chr$(Len(txtMP3(4)) + 1) & String$(3, 0) & txtMP3(4)
        
        tmpAlbum = Trim$(Mid$(strID3v1Fields, 1, 30))
        strNewTag = strNewTag & "TALB" & String$(3, 0) & Chr$(Len(tmpAlbum) + 1) & String$(3, 0) & tmpAlbum
        
        tmpYear = Trim$(Mid$(strID3v1Fields, 31, 4))
        strNewTag = strNewTag & "TYER" & String$(3, 0) & Chr$(Len(tmpYear) + 1) & String$(3, 0) & tmpYear
        
        If InStr(1, Caption, "Unregistered") > 0 Then
            strNewTag = strNewTag & "COMM" & String$(3, 0) & Chr$(29) & String$(3, 0) & "eng" & Chr$(0) & "http://LCenterprises.net"
          Else
            tmpComments = Trim$(Mid$(strID3v1Fields, 35, 30))
            strNewTag = strNewTag & "COMM" & String$(3, 0) & Chr$(Len(tmpComments) + 5) & String$(3, 0) & "eng" & Chr$(0) & tmpComments
        End If

        tmpGenre = Asc(Mid$(strID3v1Fields, 65))
        strNewTag = strNewTag & "TCON" & String$(3, 0) & Chr$(Len(tmpGenre) + 3) & String$(3, 0) & "(" & tmpGenre & ")"
        
        strNewTag = strNewTag & String$(267 - Len(strNewTag), 0)
        
        Seek tmpFile, 1

        If bolRemoveID3v1 And Left$(strID3v1, 3) = "TAG" Then 'If ID3v1 should be removed
            ReDim strFile1(LOF(tmpFile) - 129) As Byte ' Get file without Tag
            
            lblStatus = "ID3v2 does not exist, loading file..."
            DoEvents
            
            Get tmpFile, , strFile1
            Close tmpFile
        
            Open DirOpen & "\" & FilOpen For Output As tmpFile ' Erase contents of file
            Close tmpFile
        
            Open DirOpen & "\" & FilOpen For Binary As tmpFile ' Put mp3 data into file
            
            lblStatus = "ID3v2 does not exist, ID3v1 removed, rewriting file..."
            DoEvents
            Put tmpFile, , strNewTag ' put new ID3v2 Tag
            Put tmpFile, , strFile1 ' put mp3 data
          Else 'If ID3v1 should be kept or does not exist
            ReDim strFile1(LOF(tmpFile) - 1) As Byte ' get mp3 data
            
            lblStatus = "ID3v2 exists, loading file..."
            DoEvents
            
            Get tmpFile, , strFile1
            Seek tmpFile, 1
            
            lblStatus = "ID3v2 does not exist, ID3v1 kept or does not exist, rewriting file..."
            DoEvents
            Put tmpFile, , strNewTag ' put new ID3v2 Tag
            Put tmpFile, , strFile1 ' put mp3 data
        End If
    End If
    
    Close tmpFile
    
Exit Sub

Problems:
    Close tmpFile
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Sub

Private Sub cmdWriteID3v1_Click()

    WriteID3v1

End Sub

Private Sub cmdWriteID3v2_Click()

    WriteID3v2

End Sub

Private Sub DirOpen_Change()

    On Error GoTo Problems

    FilOpen = DirOpen

Exit Sub

Problems:
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Sub

Private Sub DrvOpen_Change()

    On Error GoTo Problems

    DirOpen = DrvOpen

Exit Sub

Problems:
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Sub

Private Sub FilOpen_Click()

    On Error GoTo Problems
    opID3(0).Visible = False
    opID3(1).Visible = False
    
    If ReadID3v1 = "" Then 'No ID3v1

      Else 'ID3v1 found
        opID3(0).Visible = True
        opID3(0).Value = True
    End If
        
    If ReadID3v2 = "" Then 'No ID3v2

      Else 'ID3v2 found
        opID3(1).Visible = True
        opID3(1).Value = True
    End If

Exit Sub

Problems:
    MsgBox Err.Description, vbExclamation, "Error number " & Err.Number

End Sub

Private Sub CleanTextboxes()

    For i = 1 To 4
        txtMP3(i) = ""
    Next

End Sub

Private Sub opID3_Click(Index As Integer)

    Select Case Index
      Case 0
        ReadID3v1
      Case 1
        ReadID3v2
    End Select

End Sub

':) Ulli's Code Formatter V2.0 (30.06.2001 23:53:24) 0 + 483 = 483 Lines
