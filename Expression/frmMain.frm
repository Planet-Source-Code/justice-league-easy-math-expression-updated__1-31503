VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expression"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6240
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   315
      Left            =   5280
      TabIndex        =   16
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox txtEnd 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4560
      TabIndex        =   6
      Text            =   "5"
      Top             =   1860
      Width           =   855
   End
   Begin VB.TextBox txtStart 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   4560
      TabIndex        =   4
      Text            =   "0"
      Top             =   1560
      Width           =   855
   End
   Begin VB.CheckBox chkLoop 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Loop:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3600
      TabIndex        =   13
      Top             =   960
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   315
      Left            =   2700
      TabIndex        =   11
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox txtVariables 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3600
      TabIndex        =   8
      Text            =   "A=1: B=2"
      Top             =   2460
      Width           =   2355
   End
   Begin VB.CheckBox chkVariables 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Variables:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   3600
      TabIndex        =   7
      Top             =   2160
      Value           =   1  'Checked
      Width           =   2295
   End
   Begin VB.CommandButton cmdResult 
      Caption         =   "Result"
      Default         =   -1  'True
      Height          =   315
      Left            =   3600
      TabIndex        =   12
      Top             =   3600
      Width           =   675
   End
   Begin VB.TextBox txtExpression 
      BackColor       =   &H00E0E0E0&
      Height          =   315
      Left            =   3600
      TabIndex        =   10
      Text            =   "sin(A*x)+cos(B*x)"
      Top             =   3240
      Width           =   2355
   End
   Begin VB.ListBox lstResult 
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H00000000&
      Height          =   2205
      Left            =   240
      TabIndex        =   1
      Top             =   1260
      Width           =   3195
   End
   Begin MSScriptControlCtl.ScriptControl VBSEval 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin VB.Label lblEnd 
      BackStyle       =   0  'Transparent
      Caption         =   "End:"
      Height          =   195
      Left            =   4200
      TabIndex        =   5
      Top             =   1860
      Width           =   315
   End
   Begin VB.Label lblStart 
      BackStyle       =   0  'Transparent
      Caption         =   "Start:"
      Height          =   315
      Left            =   4140
      TabIndex        =   3
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label lblResult 
      BackStyle       =   0  'Transparent
      Caption         =   "Result:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1020
      Width           =   975
   End
   Begin VB.Label lblExpression 
      BackStyle       =   0  'Transparent
      Caption         =   "&Expression:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   9
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label lblX 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Variable: X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3900
      TabIndex        =   2
      Top             =   1320
      Width           =   765
   End
   Begin MSForms.Image imgBorder 
      Height          =   3075
      Index           =   5
      Left            =   180
      Top             =   900
      Width           =   3315
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "5847;5424"
   End
   Begin MSForms.Image imgBorder 
      Height          =   3075
      Index           =   4
      Left            =   3540
      Top             =   900
      Width           =   2475
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "4366;5424"
   End
   Begin MSForms.Image imgBorder 
      Height          =   3195
      Index           =   3
      Left            =   120
      Top             =   840
      Width           =   5955
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   6
      Size            =   "10504;5636"
   End
   Begin VB.Label lblExpression 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Expression"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   1
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Width           =   5355
   End
   Begin VB.Label lblExpression 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Expression"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   2
      Left            =   375
      TabIndex        =   14
      Top             =   255
      Width           =   5355
   End
   Begin MSForms.Image imgBorder 
      Height          =   675
      Index           =   2
      Left            =   120
      Top             =   120
      Width           =   5955
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   3
      Size            =   "10504;1191"
   End
   Begin MSForms.Image imgBorder 
      Height          =   4095
      Index           =   1
      Left            =   60
      Top             =   60
      Width           =   6075
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   6
      Size            =   "10716;7223"
   End
   Begin MSForms.Image imgBorder 
      Height          =   4215
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   6195
      BackColor       =   12632256
      BorderStyle     =   0
      SpecialEffect   =   1
      Size            =   "10927;7435"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit ' all variables must be declared

'*********************************************
'           Aris J. Buenaventura
'        email: AJB2001LG@YAHOO.COM
'
'         Date Started:  February 5, 2002
'         Date Finished: February 5, 2002
'**********************************************

' 1) Click Project
' 2) Click Components...
' 3) Select Microst Script Control 1.0
' 4) Click Apply
' 5) Add ScriptControl

' if you want to make your own border
' 1) Click Project
' 2) Click Components...
' 3) Microsoft Forms 2.0 Object Library
' 4) Click Apply
' 5) Add Image
' 6) Click Image
' 7) Select Special Effect

' try this inputs
'/////////////////////////////
'   unchecked -> Loop
'   unchecked -> Variables
'   Expression:
'             1) 1+2
'             2) sin(3)
'             3) 5+cos(1)
'
'/////////////////////////////
'   unchecked -> Loop
'   Variables:
'             1) A=1
'             2) A=1: B=2
'             3) ABC=1: DEF=2
'   Expression:
'             1) sin(A)
'             2) A+B
'             3) ABC-DEF
'
'/////////////////////////////
'   Loop:
'           Value of X
'               Start: 1
'               End  : 5
'   Variables:
'           1) A=1
'           2) B=1: E=2
'           3) C=5.5
'   Expression
'           1) sin(A*x)
'           2) B+E/x
'           3) C*x

'**************************************************************************************
Private Sub Form_Load()
    Dim NewFunction As New clsAddInFunction
    
    ' Add Pi,Sec,Csc,Cot,Pow
    VBSEval.AddObject "MathFunction", NewFunction, True
    ' want more functions (see clsAddInFunction)
End Sub

'**************************************************************************************
Private Sub cmdResult_Click()
    ' expression only
    If (chkLoop.Value = vbUnchecked) And (chkVariables.Value = vbUnchecked) Then
        Expression
    End If
    
    ' expression and variables
    If (chkLoop.Value = vbUnchecked) And (chkVariables.Value = vbChecked) Then
        VariablesExpression
    End If
    
    ' expression, loop and variables
    If (chkLoop.Value = vbChecked) And (chkVariables.Value = vbChecked) Then
        LoopVariablesExpression
    End If
End Sub

'**************************************************************************************
Private Sub cmdClear_Click()
    lstResult.Clear
End Sub

'**************************************************************************************
Private Sub Expression()
    lstResult.AddItem VBSEval.Eval(txtExpression.Text) ' evaluate the expression
End Sub

'**************************************************************************************
Private Sub VariablesExpression()
    VBSEval.AddCode txtVariables.Text ' add the variables
    lstResult.AddItem VBSEval.Eval(txtExpression.Text) ' result
End Sub

'**************************************************************************************
Private Sub LoopVariablesExpression()
    Dim i As Integer
    On Error GoTo CalcErr: ' if error found then goto CalcErr
    
    For i = Val(txtStart.Text) To Val(txtEnd.Text)
        ' example -> A=1: B=2: x=[Value of i]
        VBSEval.AddCode txtVariables.Text & ": x= " & i
        lstResult.AddItem i & ") " & VBSEval.Eval(txtExpression.Text) ' result
    Next i
    Exit Sub

CalcErr:
    MsgBox Err.Description ' display error message
End Sub

'**************************************************************************************
Private Sub chkLoop_Click()
    If chkLoop.Value = vbChecked Then
        txtStart.Enabled = True
        txtEnd.Enabled = True
    Else
        txtStart.Enabled = False
        txtEnd.Enabled = False
    End If
End Sub

'**************************************************************************************
Private Sub chkVariables_Click()
    If chkVariables.Value = vbChecked Then
        txtVariables.Enabled = True
    Else
        txtVariables.Enabled = False
    End If
End Sub

'**************************************************************************************
Private Sub cmdExit_Click()
    End
End Sub
'**************************************************************************************


