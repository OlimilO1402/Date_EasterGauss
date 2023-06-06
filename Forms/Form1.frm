VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Date Of Easter-Sunday due to Gauss"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4590
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCalcEasterdate 
      Caption         =   "Easter-sunday:"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox TBYear 
      Alignment       =   1  'Rechts
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label LBEasterDate 
      Alignment       =   1  'Rechts
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   3615
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   "Year:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    'Beispiele
    '2016 war Ostern am Sonntag 27.Mrz.2016
    '1981 war Ostern am Sonntag 19.Apr.1981
    '1954 war Ostern am Sonntag 18.Apr.1954
    
    TBYear.Text = year(Now)
    LBEasterDate.Caption = "?.?." & TBYear.Text
End Sub

Private Sub BtnCalcEasterdate_Click()
    Dim y As Long: y = IIf(IsNumeric(TBYear.Text), CLng(TBYear.Text), year(Now))
    'LBEasterDate.Caption = CalcEasterdateGauss1800(y) ', ECalendar.GregorianCalendar)
    'LBEasterDate.Caption = CalcEasterdateGauss1816(y) ', ECalendar.GregorianCalendar)
    'LBEasterDate.Caption = CalcEasterdateGaussCorrected1900(y) ', ECalendar.GregorianCalendar)
    LBEasterDate.Caption = OsternShort2(y) ', ECalendar.GregorianCalendar)
End Sub


Private Sub TBYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then BtnCalcEasterdate_Click
End Sub
