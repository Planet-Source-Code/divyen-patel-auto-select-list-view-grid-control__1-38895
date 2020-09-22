VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "List View Grid Control Test Form ...."
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   11850
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8490
   ScaleWidth      =   11850
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Tip"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   80
      Width           =   735
   End
   Begin LIST_VIEW_DATA_GRID.LISTVIEW_DATAGRID LISTVIEW_DATAGRID1 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9763
      FORECOLOR       =   15951960
      AllowColumnReorder=   -1  'True
      BACKCOLOR       =   -2147483643
   End
   Begin VB.Label Label6 
      Caption         =   "Warning : When you asing query to this grid the first column must be the primary key."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4320
      TabIndex        =   5
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label5 
      Caption         =   $"Form1.frx":0442
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   7440
      Width           =   11535
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Selected Customer Id :"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   $"Form1.frx":052F
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   11655
   End
   Begin VB.Label Label1 
      Caption         =   "Customer Records"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MsgBox "Because this is created using list view control one of the advantage of this is that it automatically goes to the specified record on the bases of the first character of the key you press....", vbInformation, "Advantage ..."
End Sub

Private Sub Form_Load()
            LISTVIEW_DATAGRID1.CONNECTION_STRING = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DATABASE.MDB;Persist Security Info=False"
            LISTVIEW_DATAGRID1.OPEN_RECORD "SELECT * FROM CUSTOMERS"
            LISTVIEW_DATAGRID1.FILL_RECORDS
            LISTVIEW_DATAGRID1.SORTKEY = True
            LISTVIEW_DATAGRID1.SORT_ON_INDEX = 1
            LISTVIEW_DATAGRID1.AllowColumnReorder = True
End Sub
Private Sub LISTVIEW_DATAGRID1_GRIDCLICK()
            Label4.Caption = LISTVIEW_DATAGRID1.SELECTED_KEY
End Sub


Private Sub LISTVIEW_DATAGRID1_GRIDKEYDOWN(KEYCODE As Integer)
If KEYCODE = 13 Then
        MsgBox "SELECTED RECORD'S KEY IS : " + LISTVIEW_DATAGRID1.SELECTED_KEY
End If
End Sub
