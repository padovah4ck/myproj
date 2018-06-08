VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   ScaleHeight     =   5295
   ScaleWidth      =   6195
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ctrldeleteButton 
      Caption         =   "&Delete"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   "Deletes the "
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox ctrlTextBox1 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Text            =   "Text1"
      ToolTipText     =   "Text gets added from here to the List Box"
      Top             =   1440
      Width           =   2415
   End
   Begin VB.CommandButton ctrlAddButton 
      Caption         =   "&Add"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Adds text from the text box to the list box"
      Top             =   240
      Width           =   855
   End
   Begin VB.ListBox ctrlListBox1 
      DataField       =   "default"
      DataSource      =   "textBox1"
      Height          =   2205
      Left            =   3120
      TabIndex        =   0
      ToolTipText     =   "Items from the text box get added here"
      Top             =   360
      Width           =   2655
   End
   Begin VB.Label ctrlLabel1 
      Caption         =   "Text to add"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This method contains the code for the  click event for the add button.
Private Sub ctrlAddButton_Click()

    ' Add the string that is currently in the textBox to the List Box.
    ctrlListBox1.AddItem ctrlTextBox1.Text
    
    ' If the delete button has been disabled (because the list Box) contains
    ' no items then enable now.
    If (ctrldeleteButton.Enabled = False) Then
      ctrldeleteButton.Enabled = True
    End If
End Sub

' This method contains the code for the click event for the delete button.
Private Sub ctrlDeleteButton_Click()
    ' If the list is not empty then remove the last item from the
    ' list (adjusting the count for the number of items accordingly.  If after this remove operation
    ' the list now becomes empty, then disable the delete button.
    If (Not (ctrlListBox1.ListCount = 0)) Then
        ctrlListBox1.RemoveItem (ctrlListBox1.ListCount - 1)
        If (ctrlListBox1.ListCount = 0) Then
                ctrldeleteButton.Enabled = False
        End If
    End If
End Sub



' This method gets loaded as this whole form loads into memory.
' It checks if the List Box has no elements or not and if this is
' the case it disables the delete button (since it will otherwise
' cause an error and since it makes no sense to delete from an
' empty list).
Private Sub Form_Load()
    If (ctrlListBox1.ListCount = 0) Then
        ctrldeleteButton.Enabled = False
    End If
End Sub

