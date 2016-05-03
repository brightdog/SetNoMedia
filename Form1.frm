VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6435
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   3300
      TabIndex        =   1
      Top             =   1860
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   555
      Left            =   1800
      TabIndex        =   0
      Text            =   "e:\"
      Top             =   720
      Width           =   4515
   End
   Begin VB.Label Label2 
      Caption         =   "初始路径"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   675
      Left            =   840
      TabIndex        =   2
      Top             =   1920
      Width           =   1755
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long

Private Sub Command1_Click()
    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    i = 0
    If fso.FolderExists(Me.Text1.Text) Then
    
        SetNoMedia (Me.Text1.Text)
    
    
    Else
    
    
    End If
End Sub
Private Sub SetNoMedia(ByVal strPath As String)

    Dim fso As Scripting.FileSystemObject
    Set fso = New Scripting.FileSystemObject
    
    i = i + 1
    Me.Label1.Caption = i
    fso.CreateTextFile strPath & "\.nomedia", True
    
    
    Dim fld As Scripting.Folder
    Set fld = fso.GetFolder(strPath)
    
    Dim v As Scripting.Folder
    
    For Each v In fld.SubFolders
        DoEvents
        SetNoMedia v.Path
        
    Next

End Sub

