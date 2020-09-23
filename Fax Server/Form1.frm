VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FAX SERVER for 2000/NT"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Send with Document"
      Height          =   435
      Index           =   1
      Left            =   5265
      TabIndex        =   15
      Top             =   4365
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   435
      Index           =   0
      Left            =   5265
      TabIndex        =   14
      Top             =   3810
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1110
      Left            =   195
      TabIndex        =   9
      Top             =   3690
      Width           =   4860
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   5
         Left            =   150
         TabIndex        =   12
         Text            =   "\\SERVER"
         Top             =   615
         Width           =   1980
      End
      Begin VB.TextBox Text1 
         Height          =   300
         Index           =   4
         Left            =   2565
         TabIndex        =   10
         Text            =   "C:\Fax Covers\Basic.cov"
         Top             =   615
         Width           =   2010
      End
      Begin VB.Label Label1 
         Caption         =   "UNC Path of FAX SERVER"
         Height          =   225
         Index           =   6
         Left            =   165
         TabIndex        =   13
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "Cover Page Template"
         Height          =   225
         Index           =   5
         Left            =   2565
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox Text1 
      Height          =   1560
      Index           =   3
      Left            =   1260
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1815
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1260
      TabIndex        =   4
      Text            =   "Person, Sending Fax"
      Top             =   1410
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   1260
      TabIndex        =   2
      Text            =   "John Doe"
      Top             =   1020
      Width           =   4275
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1260
      TabIndex        =   0
      Text            =   "9,555-5555"
      Top             =   630
      Width           =   4275
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "COVER PAGE INFO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   1980
      TabIndex        =   8
      Top             =   105
      Width           =   2955
   End
   Begin VB.Label Label1 
      Caption         =   "NOTE :"
      Height          =   225
      Index           =   3
      Left            =   645
      TabIndex        =   7
      Top             =   1845
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "From :"
      Height          =   225
      Index           =   2
      Left            =   720
      TabIndex        =   5
      Top             =   1455
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Attention :"
      Height          =   225
      Index           =   1
      Left            =   435
      TabIndex        =   3
      Top             =   1065
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Fax To :"
      Height          =   225
      Index           =   0
      Left            =   585
      TabIndex        =   1
      Top             =   675
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This FAX SERVER Requires Windows 2000 or Windows NT

Private Sub Command1_Click(Index As Integer)
Dim filename1
    Select Case Index
        Case 0
        'Send the fax
            Set iFax = CreateObject("FaxServer.FaxServer")
            iFax.Connect (Text1(5).Text)
            Set IFaxDoc = iFax.CreateDocument(filename1)
            With IFaxDoc
                .FaxNumber = Text1(0).Text
                .CoverpageName = Text1(4).Text
                .SendCoverpage = True
                .RecipientName = Text1(1).Text
                .SenderName = Text1(2).Text
                .CoverpageNote = Text1(3).Text
                .Send
            End With
        Case 1
        ' Common Dialog Box
            CommonDialog1.InitDir = "C:\"
            CommonDialog1.Filter = "Office Documents (*.doc, *.rtf, *.xls, *.xlw)|*.doc;*.rtf;*.xls;*.xlw"
            CommonDialog1.DialogTitle = "Select file to Fax"
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> "" Then
                filename1 = CommonDialog1.FileName
            End If
            DoEvents
        'Send the fax "filename1"
            Set iFax = CreateObject("FaxServer.FaxServer")
            iFax.Connect (Text1(5).Text)
            Set IFaxDoc = iFax.CreateDocument(filename1)
            With IFaxDoc
                .FaxNumber = Text1(0).Text
                .FileName = filename1
                .CoverpageName = Text1(4).Text
                .SendCoverpage = True
                .RecipientName = Text1(1).Text
                .SenderName = Text1(2).Text
                .CoverpageNote = Text1(3).Text
                .Send
            End With
    End Select
Set IFaxDoc = Nothing
Set iFax = Nothing
MsgBox "The FAX has been sent!", vbOKOnly, "FAX SENT"
End Sub

Private Sub Form_Load()

End Sub
