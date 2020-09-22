VERSION 5.00
Begin VB.Form comboss 
   Caption         =   "Choose Column"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   1935
      Left            =   240
      Sorted          =   -1  'True
      Style           =   1  'Simple Combo
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
End
Attribute VB_Name = "comboss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
'MsgBox "test"

If Combo1.Text = "Boolean" Then
'frmADOX.Command3(indexss).caption = Combo1.Text
databasestructure.Command3(currentrow).caption = Combo1.Text
ElseIf Combo1.Text = "Text" Then
'frmADOX.Command3(indexss).caption = Combo1.Text
databasestructure.Command3(currentrow).caption = Combo1.Text
ElseIf Combo1.Text = "Notes" Then
'frmADOX.Command3(indexss).caption = Combo1.Text
databasestructure.Command3(currentrow).caption = Combo1.Text

ElseIf Combo1.Text = "Date" Then
'frmADOX.Command3(indexss).caption = Combo1.Text
databasestructure.Command3(currentrow).caption = Combo1.Text

ElseIf Combo1.Text = "Integer" Then
'frmADOX.Command3(indexss).caption = Combo1.Text
databasestructure.Command3(currentrow).caption = Combo1.Text
ElseIf Combo1.Text = "Currency" Then
databasestructure.Command3(currentrow).caption = Combo1.Text

Else
MsgBox "You must enter a correct column type"

Exit Sub
End If
xx = Combo1.Text

Unload Me

If xx = "Text" Then
'frmADOX.widths(indexss).Visible = True
'frmADOX.widths(indexss).SetFocus

databasestructure.widths(currentrow).Visible = True
databasestructure.widths(currentrow).Enabled = True

On Error Resume Next

databasestructure.widths(currentrow).SetFocus
On Error GoTo 0

databasestructure.check2(currentrow).Visible = False
Exit Sub
End If

If xx = "Integer" And databasestructure.check2(currentrow).Enabled = True Then

'frmADOX.check2(indexss).Visible = True
'frmADOX.check2(indexss).SetFocus
databasestructure.check2(currentrow).Visible = True
'databasestructure.widths(currentrow).Visible = False

'databasestructure.check2(currentrow).SetFocus
End If



If currentrow = databasestructure.txtFld.UBound Then
If xx = "Integer" Then
databasestructure.check2(currentrow).Visible = True
Else
databasestructure.check2(currentrow).Visible = False
End If

databasestructure.widths(currentrow).Visible = False
databasestructure.processnewrow
Else
'databasestructure.check2(currentrow).Visible = False
If xx = "Integer" Then
databasestructure.check2(currentrow).Visible = True
Else
databasestructure.check2(currentrow).Visible = False
End If

databasestructure.widths(currentrow).Visible = False
'databasestructure.processnewrow
databasestructure.txtFld(currentrow + 1).SetFocus

'frmADOX.txtFld(indexss + 1).SetFocus

End If
End If


End Sub

Private Sub Form_Load()


Combo1.AddItem "Boolean"
Combo1.AddItem "Text"
Combo1.AddItem "Notes"
Combo1.AddItem "Integer"
Combo1.AddItem "Date"
Combo1.AddItem "Currency"

'frmADOX.Visible = False
databasestructure.Visible = False

Me.Show

End Sub

Private Sub Form_Unload(Cancel As Integer)
'frmADOX.Visible = True
databasestructure.Visible = True

End Sub
