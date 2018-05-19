VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "COMPUTE"
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   2400
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1920
         TabIndex        =   1
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "GROSS"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "PERCENT"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "FICA"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1920
         Width           =   345
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim wbk As Workbook
Dim ws As Worksheet



If Trim(Text1.Text) = "" Then
    MsgBox "Please enter Gross amount.", vbCritical
    Text1.SetFocus
    Exit Sub
End If

If IsNumeric(Trim(Text1.Text)) = False Then
    MsgBox "Please enter Gross amount.", vbCritical
    Text1.SetFocus
    Exit Sub
End If

If Trim(Text2.Text) = "" Then
    MsgBox "Please enter Percent.", vbCritical
    Text2.SetFocus
    Exit Sub
End If

If IsNumeric(Trim(Text2.Text)) = False Then
    MsgBox "Please enter Percent.", vbCritical
    Text2.SetFocus
    Exit Sub
End If

Set wbk = Workbooks.Open(App.Path + "/Tax Computation Excel.xlsx", True, True)
wbk.Windows(1).Activate
wbk.Windows(1).Visible = False
Set ws = wbk.Worksheets("FICA")

    With ws
        .Cells(1, 2) = Trim(Text1.Text)
        .Cells(2, 2) = Trim(Text2.Text)
        Label3.Caption = "FICA: " & Str(.Cells(1, 4))
    End With

wbk.Close False
Set wbk = Nothing
End Sub
