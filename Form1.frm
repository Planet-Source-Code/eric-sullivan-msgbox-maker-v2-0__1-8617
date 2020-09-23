VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Eric Sullivan's Message Box Builder"
   ClientHeight    =   8970
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11025
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   11025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Caption         =   "Button Code Generator"
      Height          =   1935
      Left            =   120
      TabIndex        =   33
      Top             =   6960
      Width           =   10815
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   2280
         TabIndex        =   38
         Top             =   1080
         Width           =   8415
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   2280
         TabIndex        =   37
         Top             =   720
         Width           =   8415
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Save && Hide"
         Height          =   375
         Left            =   9240
         TabIndex        =   35
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   360
         Width           =   8415
      End
      Begin VB.Label Label15 
         Caption         =   "Label15"
         Height          =   255
         Left            =   120
         TabIndex        =   42
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label14 
         Caption         =   "Label14"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label13 
         Caption         =   "Label13"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label12 
         Caption         =   "Note: Do not change the buttons now or the code will not generate properly."
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   1490
         Width           =   6615
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2760
      TabIndex        =   24
      Top             =   6480
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Generate Code"
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   6480
      Width           =   2535
   End
   Begin VB.Frame Frame6 
      Caption         =   "Generated Code:"
      Height          =   3255
      Left            =   5400
      TabIndex        =   16
      Top             =   3600
      Width           =   5535
      Begin VB.TextBox Text6 
         Height          =   2895
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   17
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Preview:"
      Height          =   3375
      Left            =   5400
      TabIndex        =   15
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton Command9 
         Caption         =   "Generate code for buttons"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   2880
         Width           =   5055
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command6 
         Caption         =   "OK"
         Height          =   375
         Left            =   1200
         TabIndex        =   26
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Abort"
         Height          =   375
         Left            =   480
         TabIndex        =   25
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ignore"
         Height          =   375
         Left            =   3600
         TabIndex        =   22
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Retry"
         Height          =   375
         Left            =   2040
         TabIndex        =   21
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label11 
         Height          =   495
         Left            =   1080
         TabIndex        =   20
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Image Image5 
         Height          =   615
         Left            =   360
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         BackColor       =   &H00800000&
         Caption         =   " "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   285
         TabIndex        =   19
         Top             =   780
         Width           =   4950
      End
      Begin VB.Line Line4 
         X1              =   5280
         X2              =   240
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   240
         Y1              =   720
         Y2              =   2640
      End
      Begin VB.Line Line2 
         X1              =   5280
         X2              =   5280
         Y1              =   2640
         Y2              =   720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   5280
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "This is a sample of what your message box will look like."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   5295
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   1815
         Left            =   5280
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         Height          =   135
         Left            =   360
         Top             =   2640
         Width           =   5055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Options:"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   5175
      Begin VB.CheckBox Check2 
         Caption         =   "System Modal"
         Height          =   255
         Left            =   3480
         TabIndex        =   14
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1800
         TabIndex        =   13
         Top             =   360
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1800
         TabIndex        =   11
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Enter caption:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Enter window title:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Variable Name:"
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Buttons:"
      Height          =   1935
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   5175
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1200
         TabIndex        =   32
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1200
         TabIndex        =   31
         Top             =   1080
         Width           =   3735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Include help button "
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         TabIndex        =   6
         Text            =   "vbAbortRetryIgnore"
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label Label7 
         Caption         =   "Context ID:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Help File:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Icons:"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   5175
      Begin VB.Image Image6 
         Height          =   615
         Left            =   3840
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label4 
         Height          =   255
         Left            =   2400
         TabIndex        =   4
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Currently selected icon:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H000000FF&
         Height          =   615
         Left            =   3120
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "None"
         Height          =   255
         Left            =   3195
         MouseIcon       =   "Form1.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   800
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "Select the icon you wish to display in your messagebox"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   4815
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H000000FF&
         Height          =   615
         Left            =   2400
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape Shape3 
         BorderColor     =   &H000000FF&
         Height          =   615
         Left            =   1680
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         Height          =   615
         Left            =   960
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   615
         Left            =   240
         Top             =   600
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   2450
         MouseIcon       =   "Form1.frx":0152
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":02A4
         Top             =   630
         Width           =   555
      End
      Begin VB.Image Image3 
         Height          =   555
         Left            =   1680
         MouseIcon       =   "Form1.frx":06C1
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0813
         Top             =   630
         Width           =   570
      End
      Begin VB.Image Image2 
         Height          =   570
         Left            =   1000
         MouseIcon       =   "Form1.frx":0C29
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":0D7B
         Top             =   630
         Width           =   570
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   260
         MouseIcon       =   "Form1.frx":118C
         MousePointer    =   99  'Custom
         Picture         =   "Form1.frx":12DE
         Top             =   630
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TmpVar As String

Private Sub Check1_Click()
    If Check1.Value = Checked Then
        Label6.Enabled = True
        Label7.Enabled = True
        Text2.Enabled = True
        Text7.Enabled = True
    ElseIf Check1.Value = Unchecked Then
        Label6.Enabled = False
        Label7.Enabled = False
        Text2.Enabled = False
        Text7.Enabled = False
    End If
End Sub

Private Sub Combo1_Click()
    If Combo1.Text = "vbAbortRetryIgnore" Then
        Command6.Left = 1200
        Command1.Visible = True
        Command2.Visible = True
        Command3.Visible = True
        Command1.Caption = "Abort"
        Command2.Caption = "Retry"
        Command3.Caption = "Ignore"
        Command6.Visible = False
        Command7.Visible = False
    ElseIf Combo1.Text = "vbOKCancel" Then
        Command6.Left = 1200
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command6.Visible = True
        Command7.Visible = True
        Command6.Caption = "OK"
        Command7.Caption = "Cancel"
    ElseIf Combo1.Text = "vbOKOnly" Then
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command6.Visible = True
        Command6.Left = 2040
        Command6.Caption = "OK"
        Command7.Visible = False
    ElseIf Combo1.Text = "vbRetryCancel" Then
        Command6.Left = 1200
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command6.Visible = True
        Command7.Visible = True
        Command6.Caption = "Retry"
        Command7.Caption = "Cancel"
    ElseIf Combo1.Text = "vbYesNo" Then
        Command6.Left = 1200
        Command1.Visible = False
        Command2.Visible = False
        Command3.Visible = False
        Command6.Visible = True
        Command7.Visible = True
        Command6.Caption = "Yes"
        Command7.Caption = "No"
    ElseIf Combo1.Text = "vbYesNoCancel" Then
        Command6.Left = 1200
        Command1.Visible = True
        Command2.Visible = True
        Command3.Visible = True
        Command6.Visible = False
        Command7.Visible = False
        Command1.Caption = "Yes"
        Command2.Caption = "No"
        Command3.Caption = "Cancel"
    End If
End Sub

Private Sub Command4_Click()
    CreateCode
End Sub

Private Sub Command5_Click()
    End
End Sub

Private Sub Command8_Click()
    Me.Height = 7395
End Sub

Private Sub Command9_Click()
    Me.Height = 9375
    If Combo1.Text = "vbAbortRetryIgnore" Then
        Label13.Visible = True
        Label14.Visible = True
        Label15.Visible = True
        Label13.Caption = "Code for " & Command1.Caption & " button:"
        Label14.Caption = "Code for " & Command2.Caption & " button:"
        Label15.Caption = "Code for " & Command3.Caption & " button:"
        Text1.Visible = True
        Text8.Visible = True
        Text9.Visible = True
    ElseIf Combo1.Text = "vbOKCancel" Then
        Label13.Caption = "Code for " & Command6.Caption & " button:"
        Label14.Caption = "Code for " & Command7.Caption & " button:"
        Label15.Visible = False
        Text1.Visible = True
        Text8.Visible = True
        Text9.Visible = False
    ElseIf Combo1.Text = "vbOKOnly" Then
        Label13.Caption = "Code for " & Command6.Caption & " button:"
        Label14.Visible = False
        Label15.Visible = False
        Text1.Visible = True
        Text8.Visible = False
        Text9.Visible = False
    ElseIf Combo1.Text = "vbRetryCancel" Then
        Label13.Caption = "Code for " & Command6.Caption & " button:"
        Label14.Caption = "Code for " & Command7.Caption & " button:"
        Label15.Visible = False
        Text1.Visible = True
        Text8.Visible = True
        Text9.Visible = False
    ElseIf Combo1.Text = "vbYesNo" Then
        Label13.Caption = "Code for " & Command6.Caption & " button:"
        Label14.Caption = "Code for " & Command7.Caption & " button:"
        Label15.Visible = False
        Text1.Visible = True
        Text8.Visible = True
        Text9.Visible = False
    ElseIf Combo1.Text = "vbYesNoCancel" Then
        Label13.Caption = "Code for " & Command1.Caption & " button:"
        Label14.Caption = "Code for " & Command2.Caption & " button:"
        Label15.Caption = "Code for " & Command3.Caption & " button:"
        Text1.Visible = True
        Text8.Visible = True
        Text9.Visible = True
    End If
End Sub

Private Sub Form_Load()
    Me.Height = 7395
    Combo1.AddItem ("vbAbortRetryIgnore")
    Combo1.AddItem ("vbOKCancel")
    Combo1.AddItem ("vbOKOnly")
    Combo1.AddItem ("vbRetryCancel")
    Combo1.AddItem ("vbYesNo")
    Combo1.AddItem ("vbYesNoCancel")
    Command6.Visible = False
    Command7.Visible = False
    Label6.Enabled = False
    Label7.Enabled = False
    Text2.Enabled = False
    Text7.Enabled = False
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Shape1.Visible = False
    Shape2.Visible = False
    Shape3.Visible = False
    Shape4.Visible = False
    Shape5.Visible = False
End Sub

Private Sub Image1_Click()
    Label4.Caption = "Critical"
    Image5.Picture = Image1
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AlternateShapes(Shape1, Shape2, Shape3, Shape4, Shape5)
End Sub

Private Sub Image2_Click()
    Label4.Caption = "Exclamation"
    Image5.Picture = Image2
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AlternateShapes(Shape2, Shape3, Shape4, Shape5, Shape1)
End Sub

Private Sub Image3_Click()
    Label4.Caption = "Information"
    Image5.Picture = Image3
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AlternateShapes(Shape3, Shape4, Shape5, Shape1, Shape2)
End Sub

Private Sub Image4_Click()
    Label4.Caption = "Question"
    Image5.Picture = Image4
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AlternateShapes(Shape4, Shape5, Shape1, Shape2, Shape3)
End Sub

Private Sub Label2_Click()
    Label4.Caption = "No icon"
    Image5.Picture = Image6.Picture
End Sub

Private Sub label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call AlternateShapes(Shape5, Shape1, Shape2, Shape3, Shape4)
End Sub

Private Sub Text4_Change()
    Label11.Caption = Text4.Text
End Sub

Private Sub Text5_Change()
    Label10.Caption = Text5.Text
End Sub

Private Sub AlternateShapes(VisShape As Shape, NVisShape1 As Shape, NVisShape2 As Shape, NVisShape3 As Shape, NVisShape4 As Shape)
    VisShape.Visible = True
    NVisShape1.Visible = False
    NVisShape2.Visible = False
    NVisShape3.Visible = False
    NVisShape4.Visible = False
End Sub

Private Sub CreateCode()
    If Label4.Caption = "Critical" Then
        msgicon = "vbCritical"
    ElseIf Label4.Caption = "Exclamation" Then
        msgicon = "vbExclamation"
    ElseIf Label4.Caption = "Information" Then
        msgicon = "vbInformation"
    ElseIf Label4.Caption = "Question" Then
        msgicon = "vbQuestion"
    ElseIf Label4.Caption = "No icon" Then
        
    End If
    
    If Combo1.Text = "vbAbortRetryIgnore" Then
        TmpVar = vbNewLine & "   " & "Case vbAbort" & _
        vbNewLine & "   " & "   " & Text1.Text & _
        vbNewLine & "   " & "Case vbRetry" & _
        vbNewLine & "   " & "   " & Text8.Text & _
        vbNewLine & "   " & "Case vbIgnore" & _
        vbNewLine & "   " & "   " & Text9.Text
    ElseIf Combo1.Text = "vbOKCancel" Then
        TmpVar = vbNewLine & "   " & "Case vbOK" & _
        vbNewLine & "   " & "   " & Text1.Text & _
        vbNewLine & "   " & "Case vbCancel" & _
        vbNewLine & "   " & "   " & Text8.Text
    ElseIf Combo1.Text = "vbOKOnly" Then
        TmpVar = vbNewLine & "   " & "Case vbOKOnly" & _
        vbNewLine & "   " & "   " & Text1.Text
    ElseIf Combo1.Text = "vbRetryCancel" Then
        TmpVar = vbNewLine & "   " & "Case vbRetry" & _
        vbNewLine & "   " & "   " & Text1.Text & _
        vbNewLine & "   " & "Case vbCancel" & _
        vbNewLine & "   " & "   " & Text8.Text
    ElseIf Combo1.Text = "vbYesNo" Then
        TmpVar = vbNewLine & "   " & "Case vbYes" & _
        vbNewLine & "   " & "   " & Text1.Text & _
        vbNewLine & "   " & "Case vbNo" & _
        vbNewLine & "   " & "   " & Text8.Text
    ElseIf Combo1.Text = "vbYesNoCancel" Then
        TmpVar = vbNewLine & "   " & "Case vbYes" & _
        vbNewLine & "   " & "   " & Text1.Text & _
        vbNewLine & "   " & "Case vbNo" & _
        vbNewLine & "   " & "   " & Text8.Text & _
        vbNewLine & "   " & "Case vbCancel" & _
        vbNewLine & "   " & "   " & Text9.Text
    End If
    
    If Check2.Value = Checked And Check1.Value = Checked Then
        Text6.Text = Text3.Text & " = MsgBox(" & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text4.Text & Chr(34) & "," & " " & Chr(95) & _
        vbNewLine & "   " & Combo1.Text & " + vbSystemModal" & " + " & msgicon & "," & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Label10.Caption & Chr(34) & ", " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text2.Text & Chr(34) & ", " & Chr(34) & Text7.Text & Chr(34) & ")" & _
        vbNewLine & vbNewLine & "Select Case " & Text3.Text & _
        TmpVar & _
        vbNewLine & "End Select"
    ElseIf Check2.Value = Checked And Check1.Value = Unchecked Then
        Text6.Text = Text3.Text & " = MsgBox(" & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text4.Text & Chr(34) & "," & " " & Chr(95) & _
        vbNewLine & "   " & Combo1.Text & " + vbSystemModal" & " + " & msgicon & "," & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Label10.Caption & ")" & _
        vbNewLine & vbNewLine & "Select Case " & Text3.Text & _
        TmpVar & _
        vbNewLine & "End Select"
    ElseIf Check2.Value = Unchecked And Check1.Value = Checked Then
        Text6.Text = Text3.Text & " = MsgBox(" & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text4.Text & Chr(34) & "," & " " & Chr(95) & _
        vbNewLine & "   " & Combo1.Text & " + " & msgicon & "," & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Label10.Caption & Chr(34) & ", " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text2.Text & Chr(34) & ", " & Chr(34) & Text7.Text & Chr(34) & ")" & _
        vbNewLine & vbNewLine & "Select Case " & Text3.Text & _
        TmpVar & _
        vbNewLine & "End Select"
    ElseIf Check2.Value = Unchecked And Check1.Value = Unchecked Then
        Text6.Text = Text3.Text & " = MsgBox(" & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Text4.Text & Chr(34) & "," & " " & Chr(95) & _
        vbNewLine & "   " & Combo1.Text & " + " & msgicon & "," & " " & Chr(95) & _
        vbNewLine & "   " & Chr(34) & Label10.Caption & Chr(34) & ")" & _
        vbNewLine & vbNewLine & "Select Case " & Text3.Text & _
        TmpVar & _
        vbNewLine & "End Select"
    End If
End Sub
