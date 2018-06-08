VERSION 5.00
Begin VB.Form frmFirstExampleForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "My First Example"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4035
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ctrlPressMeButton 
      Caption         =   "Press Me"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      ToolTipText     =   "Go ahead press me...I know you want to."
      Top             =   240
      Width           =   1695
   End
   Begin VB.CheckBox ctrlCheckBox1 
      Caption         =   "Check1"
      Height          =   255
      Left            =   1800
      TabIndex        =   0
      ToolTipText     =   "Check me to change the backgroun"
      Top             =   2880
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Current x,y coordinates of mouse pointer"
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label ctrlCoordinateLabel 
      Caption         =   "Label1"
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      ToolTipText     =   "Displays the current x, y coordinates (doesn't work over controls though)"
      Top             =   1440
      Width           =   1815
   End
   Begin VB.Label ctrlLabel 
      Caption         =   "Check box to change background color"
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "frmFirstExampleForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem: This section would be for code outside the program

Rem: This method will determine whether or not the check box is checked.
Rem: If the check box is checked then the background color
Rem: of the form will be set to red.  If the check box is not checked then
Rem: the background color of the form will be set back to grey.
Private Sub ctrlCheckBox1_Click()

    ' If CheckBox checked then set background of form to red.
    If (ctrlCheckBox1.Value = Checked) Then
        frmFirstExampleForm.BackColor = RGB(255, 0, 0)
        
    ' If CheckBox is unchecked then set background of form to grey.
    ElseIf (ctrlCheckBox1.Value = Unchecked) Then
            frmFirstExampleForm.BackColor = RGB(210, 210, 210)
    End If
End Sub


Rem: This method will change the text displayed in the button
Rem: when the button is pressed.
Private Sub ctrlPressMeButton_Click()
    ctrlPressMeButton.Caption = "You pressed me!"
End Sub


Rem: This text box will display the current x,y coordinates of
Rem: the mouse pointer even time that the mouse moves (generating a mouse move
Rem: event). It won't update the mouse coordinates when the pointer moves over a
Rem: VB control (such as a button though).  The current x,y coordinates for the
Rem: mouse are automatically passed to this method.
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ctrlCoordinateLabel.Caption = "x = " & X & "     " & "y = " & Y
End Sub


