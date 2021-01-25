VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Form1"
   ClientHeight    =   615
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4455
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   615
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnCalcEasterdate 
      Caption         =   "Oster-Sonntag"
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
      Caption         =   "Label2"
      Height          =   255
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Jahr:"
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
Private Enum ECalendar
    JulianCalendar
    GregorianCalendar
End Enum

Private Sub Form_Load()
    'Beispiele
    '2016 war Ostern am Sonntag 27.Mrz.2016
    '1981 war Ostern am Sonntag 19.Apr.1981
    '1954 war Ostern am Sonntag 18.Apr.1954
    
    TBYear.Text = Year(Now)
    LBEasterDate.Caption = "?.?." & TBYear.Text
End Sub

Private Sub BtnCalcEasterdate_Click()
    Dim y As Long: y = IIf(IsNumeric(TBYear.Text), CLng(TBYear.Text), Year(Now))
    'LBEasterDate.Caption = CalcEasterdateGauss1800(y) ', ECalendar.GregorianCalendar)
    'LBEasterDate.Caption = CalcEasterdateGauss1816(y) ', ECalendar.GregorianCalendar)
    'LBEasterDate.Caption = CalcEasterdateGaussCorrected1900(y) ', ECalendar.GregorianCalendar)
    LBEasterDate.Caption = OsternShort2(y) ', ECalendar.GregorianCalendar)
End Sub

Private Function CalcEasterdateGauss1800(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim a As Long: a = y Mod 19 'der Mondparameter
    Dim b As Long: b = y Mod 4
    Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim N As Long
    Dim e As Long
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
    Case ECalendar.GregorianCalendar
        p = k \ 3
        q = k \ 4
        m = (15 + k - p - q) Mod 30
    End Select
    
    d = (19 * a + m) Mod 30
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        N = 6
    Case ECalendar.GregorianCalendar
        N = (4 + k - q) Mod 7
    End Select
    
    e = (2 * b + 4 * c + 6 * d + N) Mod 7
    
    OS = (22 + d + e)
    EasterMonth = 3
    If OS > 31 Then
        OS = OS - 31
        EasterMonth = 4
    End If
    Dim easter As Date: easter = OS & "." & EasterMonth & "." & y
    CalcEasterdateGauss1800 = easter
End Function

Private Function CalcEasterdateGauss1816(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim a As Long: a = y Mod 19 'der Mondparameter / Gaußsche Zykluszahl
    Dim b As Long: b = y Mod 4
    Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim N As Long
    Dim e As Long
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
    Case ECalendar.GregorianCalendar
        p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        q = k \ 4
        m = (15 + k - p - q) Mod 30
    End Select
    
    d = (19 * a + m) Mod 30
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        N = 6
    Case ECalendar.GregorianCalendar
        N = (4 + k - q) Mod 7
    End Select
    
    e = (2 * b + 4 * c + 6 * d + N) Mod 7
    
    OS = (22 + d + e)
    
    CalcEasterdateGauss1816 = CorrectOSDay(OS, y)
End Function

'Schritt     Bedeutung   Formel
'1.  die Säkularzahl                                    K(X) = X div 100
'2.  die säkulare Mondschaltung                         M(K) = 15 + (3K + 3) div 4 - (8K + 13) div 25
'3.  die säkulare Sonnenschaltung                       S(K) = 2 - (3K + 3) div 4
'4.  den Mondparameter                                  A(X) = X mod 19
'5.  den Keim für den ersten Vollmond im Frühling       D(A,M) = (19A + M) mod 30
'6.  die kalendarische Korrekturgröße                   R(D,A) = (D + A div 11) div 29[13]
'7.  die Ostergrenze                                    OG(D,R) = 21 + D - R
'8.  den ersten Sonntag im März                         SZ(X,S) = 7 - (X + X div 4 + S) mod 7
'9.  die Entfernung des Ostersonntags von der Ostergrenze
'    (Osterentfernung in Tagen)                         OE(OG,SZ) = 7 - (OG - SZ) mod 7
'10. das Datum des Ostersonntags als Märzdatum
'    (32. März = 1. April usw.)                         OS = OG + OE
Private Function CalcEasterdateGaussCorrected1900(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    Dim a As Long: a = y Mod 19 'der Mondparameter / Gaußsche Zykluszahl
    'Dim b As Long: b = y Mod 4
    'Dim c As Long: c = y Mod 7
    Dim k As Long: k = y \ 100 'die Säkularzahl
    Dim p As Long
    Dim q As Long
    Dim m As Long 'die säkulare Mondschaltung
    Dim s As Long 'die säkulare Sonnenschaltung
    Dim d As Long 'der Keim für den ersten Vollmond im Frühling
    Dim R As Long 'die kalendarische Korrekturgröße
    Dim OG As Long 'die Ostergrenze
    Dim SZ As Long 'der erste Sonntag im März
    Dim OE As Long 'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS As Long 'das Datum des Ostersonntags als Märzdatum
    Dim N As Long
    Dim e As Long
    Dim EasterMonth As Long
    
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
        s = 0
    Case ECalendar.GregorianCalendar
        p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        q = (3 * k + 3) \ 4
        m = 15 + q - p
        s = 2 - q
    End Select
    
    d = (19 * a + m) Mod 30
    R = (d + a \ 11) \ 29
    OG = 21 + d - R
    SZ = 7 - (y + y \ 4 + s) Mod 7
    OE = 7 - (OG - SZ) Mod 7
    
    OS = OG + OE
    
    CalcEasterdateGaussCorrected1900 = CorrectOSDay(OS, y)
End Function

Private Function CorrectOSDay(ByVal OS_Mrz As Long, ByVal y As Long) As Date
    Dim OSDay   As Long: OSDay = OS_Mrz + 31 * (OS_Mrz > 31)
    Dim OSMonth As Long: OSMonth = 3 - (OS_Mrz > 31)
    CorrectOSDay = DateSerial(y, OSMonth, OSDay)
End Function

Private Function OsternShort(ByVal y As Long, Optional ByVal ecal As ECalendar = ECalendar.GregorianCalendar) As Date
    'code taken from CalcEasterdateGaussCorrected1900 + CorrectOSDay
    'and then shortened
    Dim m As Long 'die säkulare Mondschaltung
    Dim s As Long 'die säkulare Sonnenschaltung
    Select Case ecal
    Case ECalendar.JulianCalendar
        m = 15
        s = 0
    Case ECalendar.GregorianCalendar
        Dim k As Long: k = y \ 100  'die Säkularzahl
        Dim p As Long: p = (8 * k + 13) \ 25 'hier unterschiedlich zu 1800
        Dim q As Long: q = (3 * k + 3) \ 4
        m = 15 + q - p
        s = 2 - q
    End Select
    
    Dim a       As Long:  a = y Mod 19                   'der Mondparameter / Gaußsche Zykluszahl
    Dim d       As Long:  d = (19 * a + m) Mod 30       'der Keim für den ersten Vollmond im Frühling
    Dim R       As Long:  R = (d + a \ 11) \ 29         'die kalendarische Korrekturgröße
    Dim OG      As Long: OG = 21 + d - R                'die Ostergrenze
    Dim SZ      As Long: SZ = 7 - (y + y \ 4 + s) Mod 7 'der erste Sonntag im März
    Dim OE      As Long: OE = 7 - (OG - SZ) Mod 7       'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS      As Long: OS = OG + OE                   'das Datum des Ostersonntags als Märzdatum
    Dim OS_Mrz  As Long: OS_Mrz = OS
    Dim OSDay   As Long: OSDay = OS_Mrz + 31 * (OS_Mrz > 31)
    Dim OSMonth As Long: OSMonth = 3 - (OS_Mrz > 31)
    OsternShort = DateSerial(y, OSMonth, OSDay)
End Function

Private Function OsternShort2(ByVal y As Long) As Date
    'let's say we only want to have GregorianCalendar
    'code taken from CalcEasterdateGaussCorrected1900 and CorrectOSDay and then shortened it
    Dim k  As Long:  k = y \ 100                                            'die Säkularzahl
                                                                            '(8 * k + 13) \ 25 'hier unterschiedlich zu 1800
    Dim q  As Long:  q = (3 * k + 3) \ 4
                                                                            '2 - q '= die säkulare Sonnenschaltung
    Dim a  As Long:  a = y Mod 19                                           'der Mondparameter / Gaußsche Zykluszahl
                                                                                      '15 + q - ((8 * k + 13) \ 25) '= die säkulare Mondschaltung
    Dim d  As Long:  d = (19 * a + (15 + q - ((8 * k + 13) \ 25))) Mod 30   'der Keim für den ersten Vollmond im Frühling
                                                                                      '(d + a \ 11) \ 29 'die kalendarische Korrekturgröße
    Dim OG As Long: OG = 21 + d - (d + a \ 11) \ 29                         'die Ostergrenze
                                                                                      '7 - (y + y \ 4 + (2 - q)) Mod 7  'der erste Sonntag im März
    Dim OE As Long: OE = 7 - (OG - (7 - (y + y \ 4 + (2 - q)) Mod 7)) Mod 7 'die Entfernung des Ostersonntags von der Ostergrenze (Osterentfernung in Tagen)
    Dim OS As Long: OS = OG + OE                                            'das Datum des Ostersonntags als Märzdatum
          OsternShort2 = DateSerial(y, (3 - (OS > 31)), (OS + 31 * (OS > 31)))
End Function

Private Sub TBYear_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then BtnCalcEasterdate_Click
End Sub
