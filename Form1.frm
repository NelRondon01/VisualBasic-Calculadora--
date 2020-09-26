VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000002&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculadora"
   ClientHeight    =   5190
   ClientLeft      =   2415
   ClientTop       =   3435
   ClientWidth     =   7215
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdOpr 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   4440
      TabIndex        =   23
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "Sqrt"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   4440
      TabIndex        =   21
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   9
      Left            =   3360
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   4440
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "Borrar"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   4440
      TabIndex        =   18
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "÷"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   6240
      TabIndex        =   17
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   6240
      TabIndex        =   16
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   6240
      TabIndex        =   15
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdOpr 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   6240
      TabIndex        =   14
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   11
      Left            =   120
      TabIndex        =   13
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   10
      Left            =   2040
      TabIndex        =   12
      Top             =   4200
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   9
      Left            =   2040
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   8
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   7
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   6
      Left            =   2040
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   5
      Left            =   1080
      TabIndex        =   7
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   3
      Left            =   2040
      TabIndex        =   5
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   1080
      TabIndex        =   4
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton CmdNum 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   0
      Left            =   1080
      TabIndex        =   2
      Top             =   4200
      Width           =   855
   End
   Begin VB.Label LblOpr 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   120
      TabIndex        =   22
      Top             =   720
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000003&
      BorderWidth     =   5
      X1              =   3120
      X2              =   3120
      Y1              =   1440
      Y2              =   4920
   End
   Begin VB.Label LblSec 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6855
   End
   Begin VB.Label LblMain 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Google Sans"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Variables Temp
Dim mainTemp As String
Dim secTemp As String
Dim opr As Integer
Dim ptTemp As Double

Private Sub CmdNum_Click(Index As Integer)
    If Index < 10 Then
        LblMain.Caption = LblMain.Caption + CStr(Index)
    ElseIf Index = 10 Then
        If LblMain.Caption <> Empty Then
            LblMain.Caption = LblMain.Caption + "00"
        End If
    ElseIf Index = 11 Then
        If LblMain.Caption <> Empty Then
            ptTemp = CInt(ptTemp) + 1
            If ptTemp <= 1 Then
                LblMain.Caption = LblMain.Caption + ","
            End If
        End If
    End If
End Sub

Private Sub CmdOpr_Click(Index As Integer)
    If LblMain.Caption <> Empty Then
        If Index = 10 Then
            LblMain.Caption = Math.Sqr(CDbl(LblMain.Caption))
        End If
    End If
    If Index = 5 And opr = 2 Then
        LblOpr.Caption = "%"
        LblMain.Caption = (CDbl(LblSec.Caption) * CDbl(LblMain.Caption)) / 100
        LblSec.Caption = Empty
    End If
    If Index = 6 Then
        LblMain.Caption = Empty
        LblSec.Caption = Empty
        LblOpr.Caption = Empty
    ElseIf Index = 4 Then
        If LblMain.Caption <> Empty Then
            LblMain.Caption = Left(LblMain.Caption, Len(LblMain.Caption) - 1)
        End If
    ElseIf Index <= 3 Then
        ptTemp = 0
        If LblMain.Caption <> Empty Then
            If Index = 0 Then
            opr = Index
            LblOpr.Caption = "+"
            mainTemp = LblMain.Caption
            secTemp = LblSec.Caption
            If LblSec.Caption = "" Then
                secTemp = 0
            End If
            LblSec.Caption = CDbl(secTemp) + CDbl(mainTemp)
            LblMain.Caption = Empty
        ElseIf Index = 1 Then
            opr = Index
            LblOpr.Caption = "-"
            mainTemp = LblMain.Caption
            secTemp = LblSec.Caption
            If LblSec.Caption = "" Then
                secTemp = 0
                LblSec.Caption = CDbl(secTemp) - CDbl(mainTemp) * -1
            Else
                LblSec.Caption = CDbl(secTemp) - CDbl(mainTemp)
            End If
            LblMain.Caption = Empty
        ElseIf Index = 2 Then
            opr = Index
            LblOpr.Caption = "x"
            mainTemp = LblMain.Caption
            secTemp = LblSec.Caption
            If LblSec.Caption = "" Then
                secTemp = 1
            End If
            LblSec.Caption = CDbl(secTemp) * CDbl(mainTemp)
            LblMain.Caption = Empty
        ElseIf Index = 3 Then
            opr = Index
            LblOpr.Caption = "÷"
            mainTemp = LblMain.Caption
            secTemp = LblSec.Caption
            If LblSec.Caption = "" Then
                secTemp = 1
                LblSec.Caption = CDbl(mainTemp) / CDbl(secTemp)
            Else
                LblSec.Caption = CDbl(secTemp) / CDbl(mainTemp)
            End If
            LblMain.Caption = Empty
        End If
        End If
    End If
    
    If Index = 9 Then
        If LblMain.Caption <> Empty And LblSec.Caption = Empty Then
            LblMain.Caption = CDbl(LblMain.Caption)
        End If
    
        If Len(LblSec.Caption) > 0 Then
            mainTemp = LblMain.Caption
            secTemp = LblSec.Caption
        
            If opr = 0 Then
                If LblMain.Caption = "" Then
                    mainTemp = 0
                End If
                LblMain.Caption = CDbl(secTemp) + CDbl(mainTemp)
            ElseIf opr = 1 Then
                If LblMain.Caption = "" Then
                    mainTemp = 0
                End If
                LblMain.Caption = CDbl(secTemp) - CDbl(mainTemp)
            ElseIf opr = 2 Then
                If LblMain.Caption = "" Then
                    mainTemp = 1
                End If
                LblMain.Caption = CDbl(secTemp) * CDbl(mainTemp)
            ElseIf opr = 3 Then
                If LblMain.Caption = "" Then
                    mainTemp = 1
                End If
                LblMain.Caption = CDbl(secTemp) / CDbl(mainTemp)
            End If
            
            LblSec.Caption = Empty
            LblOpr.Caption = Empty
        End If
    End If
End Sub

Private Sub Command1_Click()
    List1.Clear
End Sub

