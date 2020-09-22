VERSION 4.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Arrays Demo"
   ClientHeight    =   1095
   ClientLeft      =   360
   ClientTop       =   1905
   ClientWidth     =   8520
   Height          =   1560
   Icon            =   "Form1.frx":0000
   Left            =   300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Top             =   1500
   Width           =   8640
   Begin VB.CommandButton Command3 
      Caption         =   "Quit"
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It"
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   1320
      Shape           =   2  'Oval
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H000080FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   1320
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

'*****************************************************
'*                                                   *
'*                 Control Array Demo                *
'* Written By: Chris Rose                            *
'* E-Mail: Chris_1024@hotmail.com                    *
'* Description: this is just a short example that    *
'* shows you how to use control arrays               *
'* Credit: Use as you wish                           *
'*****************************************************

Dim drawn ' declare the integer drawn
Private Sub Command1_Click()



For ords = 1 To 100 'repeat 100 times
    Load Shape1(ords) 'this loads a new version of Shape1 into memory
    Shape1(ords).Visible = True 'this makes it visible
    Shape1(ords).Height = Shape1(ords - 1).Height - 2 ' this makes its height the previous shapes height - 2
    Shape1(ords).Width = Shape1(ords - 1).Height - 2 ' this does the same with its width
    Shape1(ords).Left = Shape1(ords - 1).Left + Shape1(0).Width ' this positions it next to the previous shape
Next ords ' this performs ords again

For poo = 1 To 100 'repeat 100 times
    Load Shape2(poo) 'this loads a new version of Shape2 into memory
    Shape2(poo).Visible = True 'this makes it visible
    Shape2(poo).Height = Shape2(poo - 1).Height - 2 ' this makes its height the previous shapes height - 2
    Shape2(poo).Width = Shape2(poo - 1).Height - 2 ' this does the same with its width
    Shape2(poo).Left = Shape2(poo - 1).Left + Shape2(0).Width ' this positions it next to the previous shape
Next poo ' this performs poo again

Command1.Enabled = False ' this makes command1 unselectable

drawn = True 'this sets drawn to true

Form_Resize ' this causes the form to resize

End Sub


Private Sub Command2_Click()
Unload Me 'this unloads the form and all related things
Form1.Show 'then it shows it again
End Sub

Private Sub Command3_Click()
Unload Me ' this will exit the application
End Sub

Private Sub Form_Load()
drawn = False ' this sets the integer value of drawn to false
End Sub





Private Sub Form_Resize()
If drawn = False Then ' if drawn is false then
Form1.Height = Command1.Height * 3.9 'make form1 the height of command1 * 3.9
Form1.Width = 1290 'make form1 the width of 1290
Else ' if drawn is true then
Form1.Width = Command1.Width + (Shape1(0).Width * 102) / 1.8 + 200 'make form1's width the right size to fit in all the squares
End If 'close the if then clause
End Sub


