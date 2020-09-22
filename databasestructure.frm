VERSION 5.00
Begin VB.Form databasestructure 
   Caption         =   "Change Database Structure"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5220
   ScaleWidth      =   7200
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Function Keys"
      Height          =   255
      Left            =   1320
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   2760
      Width           =   1335
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2400
      Width           =   4215
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1695
      Left            =   6240
      TabIndex        =   9
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   600
      Width           =   6135
      Begin VB.CheckBox check2 
         Caption         =   "Primary Key"
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox txtFld 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox widths 
         Height          =   285
         Index           =   0
         Left            =   3960
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Height          =   285
         Index           =   0
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1815
      End
      Begin VB.CheckBox chkrequired 
         Caption         =   "Required"
         Height          =   285
         Index           =   0
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Name Of Field"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Column Type"
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Width/Primary"
      Height          =   255
      Left            =   3960
      TabIndex        =   11
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Table Name"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "databasestructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hscrollvals As Integer
Private dirty As Boolean
Private hasentered As Boolean
Private Type columninfo
'isnew As Boolean
    oldname As String
    oldwidth As Integer
End Type
Private oldtable As String
Private deleterow() As columninfo
Private cat As New ADOX.Catalog


Private numtexts As Integer
Private doscroll As Boolean

Private colinfo() As columninfo
Public Function ConvType(ByVal TypeVal As Long) As String
'Combo1.AddItem "Boolean"
'Combo1.AddItem "Text"
'Combo1.AddItem "Notes"
'Combo1.AddItem "Integer"
'Combo1.AddItem "Date"
  Select Case TypeVal
        Case adBigInt                    ' 20
            ConvType = "Big Integer"
        Case adBinary                    ' 128
            ConvType = "Binary"
        Case adBoolean                   ' 11
            ConvType = "Boolean"
        Case adBSTR                      ' 8 i.e. null terminated string
            ConvType = "Text"
        Case adChar                      ' 129
            ConvType = "Text"
        Case adCurrency                  ' 6
            ConvType = "Currency"
        Case adDate                      ' 7
            ConvType = "Date"
        Case adDBDate                    ' 133
            ConvType = "Date"
        Case adDBTime                    ' 134
            ConvType = "Date"
        Case adDBTimeStamp               ' 135
            ConvType = "Date"
        Case adDecimal                   ' 14
            ConvType = "Float"
        Case adDouble                    ' 5
            ConvType = "Float"
        Case adEmpty                     ' 0
            ConvType = "Empty"
        Case adError                     ' 10
            ConvType = "Error"
        Case adGUID                      ' 72
            ConvType = "GUID"
        Case adIDispatch                 ' 9
            ConvType = "IDispatch"
        Case adInteger                   ' 3
            ConvType = "Integer"
        Case adIUnknown                  ' 13
            ConvType = "Unknown"
        Case adLongVarBinary             ' 205
            ConvType = "Binary"
        Case adLongVarChar               ' 201
            ConvType = "Notes"
        Case adLongVarWChar              ' 203
            ConvType = "Notes"
        Case adNumeric                  ' 131
            ConvType = "Long"
        Case adSingle                    ' 4
            ConvType = "Single"
        Case adSmallInt                  ' 2
            ConvType = "Small Integer"
        Case adTinyInt                   ' 16
            ConvType = "Tiny Integer"
        Case adUnsignedBigInt            ' 21
            ConvType = "Big Integer"
        Case adUnsignedInt               ' 19
            ConvType = "Integer"
        Case adUnsignedSmallInt          ' 18
            ConvType = "Small Integer"
        Case adUnsignedTinyInt           ' 17
            ConvType = "Timy Integer"
        Case adUserDefined               ' 132
            ConvType = "UserDefined"
        Case adVarNumeric                 ' 139
            ConvType = "Long"
        Case adVarBinary                 ' 204
            ConvType = "Binary"
        Case adVarChar                   ' 200
            ConvType = "Text"
        Case adVariant                   ' 12
            ConvType = "Variant"
        Case adVarWChar                  ' 202
            ConvType = "Text"
        Case adWChar                     ' 130
            ConvType = "Text"
        Case Else
            ConvType = "Unknown"
   End Select
End Function





Private Sub check2_GotFocus(Index As Integer)
Dim newscrolls As Integer
If HScroll1.Max > 0 Then
newscrolls = Command3(0).Left + Command3(0).Width
newscrolls = newscrolls / 100
newscrolls = newscrolls - 2
HScroll1.Value = newscrolls
End If
currentrow = Index

End Sub

Private Sub check2_LostFocus(Index As Integer)
'If index > 0 And KeyCode = 38 Then
'If check2(index - 1).Visible = True Then

'set focus up

'dirty = True
'Exit Sub
'End If

'ElseIf KeyCode = 40 And txtFld.UBound > 0 Then
'If check2(index + 1).Visible = True Then
'set focus down

'dirty = True
'Exit Sub
'End If


detprocess Index, False

'End If
End Sub

Private Sub chkrequired_GotFocus(Index As Integer)
Dim newscrolls As Integer
If HScroll1.Max > 0 Then
newscrolls = widths(0).Left + widths(0).Width
newscrolls = newscrolls / 100
newscrolls = newscrolls - 2
If newscrolls > HScroll1.Max Then
HScroll1.Value = HScroll1.Max
Else
HScroll1.Value = newscrolls

End If

'HScroll1.Value = newscrolls
End If
currentrow = Index

End Sub

Private Sub Command1_Click()
    'savechanges
    MySaveChanges
End Sub

Private Sub Command2_Click()
MsgBox "F1 to save table information" & Chr(10) & "F2 to delete current row" & Chr(10) & "F3 to mark/unmark field required" & Chr(10) & "F4 to exit without saving table information" & Chr(10) & "F5 to delete table"
Text1.SetFocus

End Sub

Private Sub Command3_Click(Index As Integer)
Load comboss
End Sub

Private Sub Command3_GotFocus(Index As Integer)
Dim newscrolls As Integer

If HScroll1.Max > 0 Then
'HScroll1.Value = txtFld(0).Width / 100
newscrolls = txtFld(0).Width / 100
If newscrolls > HScroll1.Max Then
HScroll1.Value = HScroll1.Max
Else

HScroll1.Value = newscrolls - 2
End If

End If
currentrow = Index

End Sub
Private Sub savechanges()
'MsgBox "changes saved"
Dim x As Integer

    If currentrow = -1 Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If
    
    If txtFld(currentrow).Text <> "" And Command3(currentrow).caption = "" Then
        MsgBox "Sorry, you must choose a column type"
        Load comboss
        Exit Sub
    End If
    For x = 0 To txtFld.UBound
    
        If txtFld(x).Text = txtFld(Index) And x <> Index Then
            MsgBox "You must enter a unique column name"
            txtFld(Index).SetFocus
            Exit Sub
            Exit For
        End If
    Next

    If txtFld(currentrow).Text = "" And colinfo(currentrow).oldname <> "" Then
        MsgBox "You cannot change the name of the column to blank"
        txtFld(currentrow).SetFocus
        Exit Sub
    End If

    If oldtable = "" Then
        'MsgBox "create new table"
        Set tblnew = New ADOX.Table
            tblnew.Name = Text1.Text
         Set tblnew.ParentCatalog = cat
         cat.Tables.Append tblnew
         cat.Tables.Refresh
     
    End If
    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable <> "" Then
        'MsgBox "will delete table"
        sqls = "drop table " & oldtable
        dbObj.Execute (sqls)
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable = "" Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If oldtable <> "" Then

        If UBound(deleterow) > 0 Then

            For x = 1 To UBound(deleterow)
                'MsgBox deleterow(x).oldname
                sqls = "alter table " & oldtable & " drop column " & deleterow(x).oldname
                dbObj.Execute sqls
            Next
        End If
    End If


'MsgBox "changes saved"

'Dim catdb As New ADOX.Catalog

'Set catdb = New ADOX.Catalog
    'catDB.Create CnnString
    'catdb.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\Personal Information\Test1.mdb;Persist Security Info=False"


'For x = 1 To 30
'MsgBox "test"
    If oldtable <> Text1.Text And oldtable <> "" Then
        cat.Tables.Item(oldtable).Name = Text1.Text
        cat.Tables.Refresh
    End If

    For x = 0 To txtFld.UBound
        'renaming columns
        If txtFld(x).Text <> "" Then
            If colinfo(x).oldname <> "" And colinfo(x).oldname <> txtFld(x).Text Then
                cat.Tables.Item(Trim(Text1.Text)).Columns.Item(colinfo(x).oldname).Name = txtFld(x).Text
                cat.Tables.Item(Trim(Text1.Text)).Columns.Refresh
                cat.Tables.Refresh
            Else
                ColName = Trim(Me.txtFld(x).Text)

    Select Case LCase(Trim(Command3(x).caption))
        Case "Text"
            colType = adVarWChar   'adVarChar
        'Case "float"
            'colType = adVarNumeric
        Case "Integer"
            colType = adInteger
        Case "Date"
            colType = adDate
        Case "boolean"
            colType = adBoolean
        Case "Notes"
            colType = adLongVarWChar
        Case "Currency"
            colType = adCurrency
    End Select
    
                 If Command3(x).caption = "Text" And widths(x).Text = "" Then
                     ColWidth = 50
                 Else
                     ColWidth = IIf(Trim(Me.widths(x).Text) = "", 0, widths(x).Text)
                 End If
                 
                'hidden bug (only affects when new table is created)
       
                cat.Tables.Item(Text1.Text).Columns.Append ColName, colType, ColWidth
                
                'If Me.chkPrimary.Value = 1 Then
                    'idx.Name = ColName   'replace with new name for index
                    'idx.Columns.Append ColName  'column to be primary key
                    'idx.PrimaryKey = True       'set as primary key
                    'idx.Unique = True           'set as unique
                    'cat.Tables.Item(sTableName).Indexes.Append idx  'add that index into current table
                'End If
    
                'auto increment field
                If check2(x).Value = 1 Then
                    cat.Tables.Item(Text1.Text).Columns.Item(ColName).Properties("AutoIncrement") = True
                End If
                If chkrequired(x).Value = 0 Then
                'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
                'cat.Tables.Item(Text1.Text).Columns.Item(txtFld(x).Text).Attributes = adColNullable
                'On Error Resume Next
                    cat.Tables.Item(Text1.Text).Columns.Item(txtFld(x).Text).Attributes = adColNullable
                'On Error GoTo 0
                End If
    
                cat.Tables.Item(Text1.Text).Columns.Refresh
                cat.Tables.Refresh
                
                'cat.Tables.Item("test2").Columns.Refresh
                'cat.Tables.Refresh
            End If
        End If
    Next
    Unload Me
    Call frmMain.SetupDatabase(Text1.Text)
    Exit Sub
'End If


'MsgBox "create new table"



End Sub
Private Sub deleterows()
'MsgBox "test"



Dim tmpCN As New ADODB.Connection   'temporary connection
    Set tmpCN = New ADODB.Connection
    
    Dim scmd As String
    
    
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName         'change tmpColumn with a variable - later
    tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
'MsgBox "row " & currentrow & " has been deleted"
If currentrow = -1 Then
MsgBox "Sorry, cannot delete row because you are not on a row"
Exit Sub
End If
If currentrow = 1 And txtFld.UBound = 1 And oldtable = "" Then

'MsgBox "Will delete the contents of row 0 only"
txtFld(1).Text = ""
Command3(1).caption = ""
chkrequired(1).Value = 0
check2(1).Value = 0
widths(1).Text = ""
widths(1).Visible = False
check2(1).Visible = False
Command3(1).Enabled = True

txtFld(1).SetFocus

'If oldtable <> "" Then
'MsgBox "Will delete contents of table"
'End If
Exit Sub
End If
If oldtable <> "" Then
Dim nquestion As Integer
nquestion = MsgBox("Are you sure you want to delete the row because if you do, you will lose the data as well", vbYesNo)
If nquestion = 7 Then
Exit Sub
End If

'ask1 = MsgBox("Are you sure you want to do this move.  If you do, you will knock yourself out", vbYesNo)




'ReDim Preserve deleterow(UBound(deleterow) + 1)
'deleterow(UBound(deleterow)).oldname = colinfo(currentrow).oldname


sqls = "alter table " & oldtable & " drop column " & colinfo(currentrow).oldname
tmpCN.Execute (sqls)
'On Error Resume Next
cat.Tables(oldtable).Columns.Refresh
'On Error GoTo 0
cat.Tables.Refresh

'MsgBox currentrow
Dim deletetables As Boolean
deletetables = False

If currentrow = 1 And check2(0).Value = 1 And currentrow = txtFld.UBound Or currentrow = 0 And currentrow = txtFld.UBound Then
sqls = "drop table " & oldtable
tmpCN.Execute (sqls)
deletetables = True
End If


'If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable <> "" Then
        'MsgBox "will delete table"
        'sqls = "drop table " & oldtable
        'dbObj.Execute (sqls)
        'tmpCN.Execute (sqls)
        'Call frmMain.SetupDatabase(Text1.Text)
        'Unload Me
        
        'Exit Sub
    'End If
    
    'if txtfld.UBound=0 and check2(0).Value=1 and oldtable<>""


'Set cat = New ADOX.Catalog
'cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath'
'tmpCN.Close
'Set tmpCN = Nothing
'If currentrow <> 1 And check2(0).Value = 0 And currentrow = txtFld.UBound or currentrow


If deletetables = False Then


Call frmMain.SetupDatabase(Text1.Text)
Unload Me
'frmMain.mnuedittable_Click
frmMain.newloads
Else
Unload Me
Call frmMain.SetupDatabase("")
End If

'next, transfer values
'then delete row from array and controls
'form_resize




Exit Sub





End If


Dim x As Integer
For x = currentrow + 1 To txtFld.UBound
txtFld(x - 1) = txtFld(x)
Command3(x - 1) = Command3(x)
check2(x - 1) = check2(x)
chkrequired(x - 1) = chkrequired(x)
widths(x - 1) = widths(x)
colinfo(x - 1) = colinfo(x)

Next
Unload txtFld(txtFld.UBound)
Unload Command3(Command3.UBound)
Unload check2(check2.UBound)
Unload widths(widths.UBound)
Unload chkrequired(chkrequired.UBound)
ReDim Preserve colinfo(UBound(colinfo) - 1)
Form_Resize
'MsgBox currentrow
txtFld_GotFocus (currentrow)

'txtFld(currentrow).SetFocus

'MsgBox "Will delete row " & currentrow



End Sub
Private Sub checkprimary()
If currentrow = -1 Then
MsgBox "You must be on a row in order to checkbox primary key"
Exit Sub
End If
If colinfo(currentrow).oldname <> "" Then
MsgBox "Sorry, you cannot select/unselect primary because this field is not new"
Exit Sub
End If

If check2(currentrow).Visible = False Then
MsgBox "You cannot select/unselect primary because the checkbox is not visible"
Exit Sub
End If


If check2(currentrow).Value = 1 Then
check2(currentrow).Value = 0
ElseIf check2(currentrow).Value = 0 Then
check2(currentrow).Value = 1
End If

'MsgBox "row " & currentrow & " auto increment"

End Sub
Private Sub checkrequired()
If currentrow = -1 Then
MsgBox "You must be on a row in order to checkbox required"
Exit Sub
End If
If colinfo(currentrow).oldname <> "" Then
MsgBox "Sorry, you cannot select/unselect required because it is not a new column"
Exit Sub
End If
If oldtable <> "" Then
MsgBox "Sorry, you cannot make this a required field since you are adding a column"
Exit Sub
End If



If chkrequired(currentrow).Value = 1 Then
chkrequired(currentrow).Value = 0
ElseIf chkrequired(currentrow).Value = 0 Then
chkrequired(currentrow).Value = 1
End If


'MsgBox "row " & currentrow & " checkrequired"

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then
        'MsgBox "test 1"
        
        dirty = True
        MySaveChanges
    ElseIf KeyCode = 113 Then
        'MsgBox "test 2"
        
        dirty = True
        deleterows
    ElseIf KeyCode = 114 Then
        'dirty = True
        'MsgBox "test 3"
        
        checkrequired
    'ElseIf KeyCode = 115 Then
        'dirty = True
        'MsgBox "test 4"
        
        'checkprimary
    
    ElseIf KeyCode = 115 Then
        'MsgBox "test 4"
        
        dirty = True
        Unload Me
    ElseIf KeyCode = 116 Then
        'MsgBox "test 5"
    
        If oldtable = "" Then
            MsgBox "Sorry, no table exists to drop"
            Exit Sub
        End If
        
        
        dirty = True
        sqls = "drop table " & oldtable
        dbObj.Execute (sqls)
        Unload Me
        Call frmMain.SetupDatabase("")
    
        
    End If

End Sub

Private Sub Form_Load()

'changes
'autoincrement automatically on all new databases
'if new database, cannot change autoincrement
'on existing, cannot do anything about it
'cannot delete rows for autoincrements
'cannot use name id for column names
'cannot checkbox autoincrement
'autoincrement will show up as first line on all new databases
'will also be disabled
'if disabled, then go up 1 more or go down 1 more




currentrow = -1
ReDim deleterow(0)
Set cat = New ADOX.Catalog
cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
If isnewtable = False Then
'MsgBox "test"

'F1 on form will save changes
'F2 on form will delete current row if possible
'F3 on form will check/uncheck the required
'F4 on form will check/uncheck the primary key
'MsgBox "test"


            oldtable = frmMain.cmbTables.Text
            Text1.Text = oldtable
            
            For nCount = 0 To cat.Tables.Item(oldtable).Columns.Count - 1
            If nCount > 0 Then
            Load txtFld(nCount)
            Set txtFld(nCount).Container = Frame1
            txtFld(nCount).Top = txtFld(nCount - 1).Top + txtFld(nCount - 1).Height
            txtFld(nCount).Visible = True
            txtFld(nCount).Enabled = True
            Load widths(nCount)
            Set widths(nCount).Container = Frame1
            widths(nCount).Top = txtFld(nCount - 1).Top + txtFld(nCount - 1).Height
            'widths(nCount).Visible = False
            widths(nCount).Enabled = True
            'txtFld(nCount).Visible = True
            Load Command3(nCount)
            Set Command3(nCount).Container = Frame1
            'Command3(nCount).caption = ""
            
            Command3(nCount).Top = txtFld(nCount - 1).Top + txtFld(nCount - 1).Height
            Command3(nCount).Visible = True
            'Command3(nCount).Enabled = False
            Load check2(nCount)
            Set check2(nCount).Container = Frame1
            check2(nCount).Top = txtFld(nCount - 1).Top + txtFld(nCount - 1).Height
            'check2(nCount).Visible = False
            Load chkrequired(nCount)
            Set chkrequired(nCount).Container = Frame1
            chkrequired(nCount).Top = txtFld(nCount - 1).Top + txtFld(nCount - 1).Height
            chkrequired(nCount).Visible = True
            Else
            txtFld(nCount).Top = 0
            chkrequired(nCount).Top = 0
            Command3(nCount).Top = 0
            widths(nCount).Top = 0
            check2(ncont).Top = 0
            
            End If
            'Command3(nCount).caption = frmViewColumns.ConvType(frmMain.adoData.Recordset.Fields(nCount).Type)
            'Command3(nCount).caption = frmViewColumns.ConvType(cat.Tables(oldtable).Columns(nCount).Type)
            Command3(nCount).caption = ConvType(cat.Tables(oldtable).Columns(nCount).Type)
            
            '*************************************
            'Command3(nCount).Enabled = False
            Command3(nCount).Enabled = True
            
            
            check2(nCount).Visible = False
            'widths(nCount).Visible = False
            ReDim Preserve colinfo(nCount)
            
            txtFld(nCount).Text = cat.Tables(oldtable).Columns(nCount).Name
            
            
            
            'txtFld(nCount).Text = frmMain.adoData.Recordset.Fields(nCount).Name
            
            'colinfo(nCount) = txtFld(nCount).Text
            colinfo(nCount).oldname = txtFld(nCount).Text
            'MsgBox Command3(nCount).caption
            'MsgBox "Test"
            
            If Command3(nCount).caption = "Text" Then
            'MsgBox nCount
            
            'colinfo(nCount).oldwidth = frmMain.adoData.Recordset.Fields(nCount).DefinedSize
            'widths(nCount).Text = colinfo(nCount).oldwidth
            widths(nCount).Text = cat.Tables(oldtable).Columns(nCount).DefinedSize
            colinfo(nCount).oldwidth = widths(nCount).Text
            'MsgBox widths(nCount).Text
            
            widths(nCount).Visible = True
            'Else
            'widths(nCount).Visible = False
            End If
            
            If Command3(nCount).caption = "Integer" And cat.Tables(oldtable).Columns(nCount).Properties("AutoIncrement") = True Then
            'MsgBox txtFld(nCount).Text & "  " & nCount
            
            
            check2(nCount).Value = 1
            check2(nCount).Visible = True
            ElseIf Command3(nCount).caption = "Integer" Then
            check2(nCount).Visible = True
            check2(nCount).Value = 0
            Else
            check2(nCount).Visible = False
            check2(nCount).Value = 0
            End If
            If cat.Tables(oldtable).Columns(nCount).Attributes = adColNullable Then
            chkrequired(nCount).Value = 0
            Else
            chkrequired(nCount).Value = 1
            End If
            
            'check2(nCount).Visible = True
            
            '******************************
            'widths(nCount).Enabled = False
            widths(nCount).Enabled = True
            'widths(nCount).Visible = True
            'check2(nCount).Visible = False
            
            chkrequired(nCount).Enabled = False
            'check2(nCount).Enabled = False
            'MsgBox "test"
            
            If check2(nCount).Value = 1 Then
            'MsgBox check2(nCount).Top
            
            txtFld(nCount).Enabled = False
            check2(nCount).Enabled = False
            widths(nCount).Enabled = False
            Command3(nCount).Enabled = False
            chkrequired(nCount).Enabled = False
            'Else
            'txtFld(nCount).Enabled = True
            'widths(nCount).Enabled = True
            
            End If
            Command3(nCount).Enabled = False
            check2(nCount).Enabled = False
            
            'End If
            
                '.TextMatrix(nCount, 0) = frmMain.adoData.Recordset.Fields(nCount - 1).Name
                '.TextMatrix(nCount, 1) = ConvType(frmMain.adoData.Recordset.Fields(nCount - 1).Type)
                '.TextMatrix(nCount, 2) = frmMain.adoData.Recordset.Fields(nCount - 1).DefinedSize
                '.TextMatrix(ncount, 3)=frmmain.adoData.Recordset.Fields(ncount-1).
            Next nCount
        'End If
    'End With





End If
'MsgBox "test"

If isnewtable = True Then
currentrow = 0

ReDim colinfo(0)
colinfo(0).oldname = ""
oldtable = ""
'colinfo(1).oldname = ""
txtFld(0).Enabled = False
chkrequired(0).Enabled = False
widths(0).Enabled = False
check2(0).Enabled = False
check2(0).Visible = True
Command3(0).Enabled = False
txtFld(0).Text = "id"
Command3(0).caption = "Integer"
chkrequired(0).Value = 1
check2(0).Value = 1
'Load txtFld(1)
processnewrow

'Else
'change later
'ReDim colinfo(0)
'colinfo(0) = ""

End If


HScroll1.SmallChange = 2
HScroll1.LargeChange = 10

'Form_Resize
VScroll1.TabStop = False
HScroll1.TabStop = False
Command1.TabStop = False
Me.Show
'MsgBox check2(0).Top & "  " & check2(0).Value & "  " & txtFld(0).Top



End Sub
Public Sub processnewrow()
'MsgBox "new row" & "  " & currentrow + 1 & " row number next"
'MsgBox "test"

Load txtFld(currentrow + 1)
Set txtFld(txtFld.UBound).Container = Frame1
txtFld(txtFld.UBound).Top = txtFld(currentrow).Top + txtFld(currentrow).Height
txtFld(txtFld.UBound).Visible = True
txtFld(txtFld.UBound).Enabled = True
txtFld(txtFld.UBound).Text = ""


Load Command3(currentrow + 1)
Set Command3(Command3.UBound).Container = Frame1
Command3(Command3.UBound).Enabled = True
Command3(Command3.UBound).Top = txtFld(currentrow).Top + txtFld(currentrow).Height
Command3(Command3.UBound).Visible = True
Command3(Command3.UBound).caption = ""

Load check2(currentrow + 1)
Set check2(check2.UBound).Container = Frame1
check2(check2.UBound).Visible = False
'check2(check2.UBound).Enabled = True

check2(check2.UBound).Top = txtFld(currentrow).Top + txtFld(currentrow).Height
check2(check2.UBound).Value = 0

Load widths(currentrow + 1)
Set widths(widths.UBound).Container = Frame1
widths(widths.UBound).Top = txtFld(currentrow).Top + txtFld(currentrow).Height
widths(widths.UBound).Visible = False
widths(widths.UBound).Text = 50

Load chkrequired(currentrow + 1)
Set chkrequired(chkrequired.UBound).Container = Frame1
chkrequired(chkrequired.UBound).Top = txtFld(currentrow).Top + txtFld(currentrow).Height

chkrequired(chkrequired.UBound).Visible = True
If oldtable = "" Then

chkrequired(chkrequired.UBound).Enabled = True
End If

chkrequired(chkrequired.UBound).Value = 0

'Form_Resize
On Error Resume Next

txtFld(txtFld.UBound).SetFocus
ReDim Preserve colinfo(UBound(colinfo) + 1)
colinfo(UBound(colinfo)).oldname = ""
On Error GoTo 0



End Sub

Private Sub Form_Resize()
On Error Resume Next

Dim newx As Integer

Dim scrolls As Integer
Frame1.Height = Me.Height - Text1.Height - Command1.Height - 900 - HScroll1.Height

Frame1.Width = Me.Width - VScroll1.Width - 200
'HScroll1.Top = Frame1.Height - HScroll1.Height - 400
HScroll1.Top = Frame1.Height + Frame1.Top

'VScroll1.Left = Frame1.Width - VScroll1.Width - 400
VScroll1.Left = Frame1.Left + Frame1.Width

HScroll1.Left = 50
'VScroll1.Top = 50
VScroll1.Top = Frame1.Top


HScroll1.Width = Frame1.Width
VScroll1.Height = Frame1.Height

Command1.Top = Frame1.Top + Frame1.Height + HScroll1.Height + 50
Command2.Top = Command1.Top
Command2.Left = Command1.Left + Command1.Width

If Frame1.Width < chkrequired(0).Width + chkrequired(0).Left + 200 And txtFld(0).Left = 0 Then

scrolls = chkrequired(x).Width + chkrequired(0).Left + 200 - Frame1.Width
scrolls = scrolls / 100
'MsgBox scrolls
'each number, add 100 or subtract 100

HScroll1.Max = scrolls
hscrollvals = HScroll1.Max

'MsgBox HScroll1.Max

ElseIf txtFld(0).Left = 0 Then

HScroll1.Max = 0

End If
'change later

newx = VScroll1.Height / txtFld(0).Height
searches = InStr(newx, ".")
If searches > 0 Then

newx = Mid(newx, 1, searches - 1)
End If

'MsgBox newx
'if txtfld(txtfld.UBound).Top+txtfld(txtfld(ubound).Height>
'MsgBox "test"

'If txtFld(txtFld.ubound).Top + txtFld(txtFld.ubound).Height > VScroll1.Height + VScroll1.Top - Label3.Height - 75 Then
If txtFld(txtFld.UBound).Top + txtFld(txtFld.UBound).Height > VScroll1.Height + VScroll1.Top - Label3.Height - 300 Then









'if txtfld(txtfld(ubound)).Top="" then

'If newx - 1 < txtFld.UBound Then
VScroll1.Max = txtFld.UBound
'If newx > 2 Then
On Error Resume Next

VScroll1.SmallChange = 2
VScroll1.LargeChange = newx - 2
On Error GoTo 0

'Else
'VScroll1.SmallChange = 1
'VScroll1.LargeChange = 1
'End If

Else

VScroll1.Max = 0
End If
On Error GoTo 0

Exit Sub


'newx = VScroll.Height / Text1(1).Height
'searches = InStr(newx, ".")

'newx = Mid(newx, 1, searches - 1)


'If VScroll.Value = 0 Then
'VScroll.LargeChange = newx
'Else


'VScroll.LargeChange = newx - 1
'End If

'VScroll.SmallChange = newx / 3
'If Text1(texts).Top + Text1(texts).Height > VScroll.Top + VScroll.Height Then


'VScroll.Max = texts


'End If


End Sub

Private Sub HScroll1_Change()
Dim newposition As Long
If HScroll1.Value = 0 Then
newposition = 0
Else
newposition = HScroll1.Value - 1
newposition = newposition * CLng(-100)
End If

'For x = 1 To texts
'MsgBox Text1(x).Text

'Text1(x).Top = (VScroll.Value * -100) + xx
'MsgBox Text1(x)
'MsgBox Text1(x).Top
'On Error GoTo messages
'newx = CLng(VScroll.Value) * CLng(-200)
'newx = newx + CLng(xx)
Dim newwidths As Integer

For x = 1 To 4
'3
'2
'4
''

If x = 1 Then
newwidths = txtFld(0).Width
ElseIf x = 2 Then
newwidths = Command3(0).Width
ElseIf x = 3 Then
newwidths = widths(0).Width
ElseIf x = 4 Then
newwidths = chkrequired(0).Width
End If
If x = 1 Then

Label3.Left = newposition
ElseIf x = 2 Then
Label2.Left = newposition
ElseIf x = 3 Then
Label4.Left = newposition
End If


For y = 0 To txtFld.UBound
If x = 1 Then
txtFld(y).Left = newposition
ElseIf x = 2 Then
Command3(y).Left = newposition
ElseIf x = 3 Then
widths(y).Left = newposition
check2(y).Left = newposition
ElseIf x = 4 Then
chkrequired(y).Left = newposition
End If


Next
newposition = newposition + newwidths

Next

'Text1(x).Top = newposition
'newposition = newposition + txtFld(0).Width


'xx = xx + 300
'Next
'If Text1(texts).Top + Text1(texts).Height < VScroll.Top + VScroll.Height Then
If Frame1.Width > chkrequired(0).Left + chkrequired(0).Width + 200 Then
'And txtFld(0).Left = 0

'MsgBox Text1(100).Top
'MsgBox Form1.Height + StatusBar1.Height
'MsgBox VScroll.Height
'MsgBox VScroll.Top

HScroll1.Max = HScroll1
Else

HScroll1.Max = hscrollvals



End If



End Sub
Private Sub updatescroll(Index As Integer)
If VScroll1.Max > 0 And Index < 2 Then
VScroll1.Value = 0
'ElseIf VScroll1.Value < VScroll1.Max Then
Else

newscroll = VScroll1.Value + (Index - currentrow)
If newscroll < 0 Then
VScroll1.Value = 0
ElseIf newscroll < VScroll1.Max Then
'MsgBox newscroll

If newscroll < 0 Then
newscroll = newscroll - 1
ElseIf VScroll1.Value = 0 Then
VScroll1.Value = 2
Else

'newscroll = newscroll

VScroll1.Value = newscroll
End If
End If



End If
End Sub

Private Sub Text1_GotFocus()
On Error Resume Next

With Screen.ActiveForm
.ActiveControl.SetFocus
.ActiveControl.SelStart = 0
.ActiveControl.SelLength = Len(.ActiveControl.Text)
End With
currentrow = -1

End Sub

Private Sub Text1_LostFocus()
'txtFld(0).SetFocus
'MsgBox "test"
If Text1.Text = "" Then
Text1.SetFocus
MsgBox "You must enter the table name"
End If

End Sub

Private Sub txtFld_Change(Index As Integer)
hasentered = True

End Sub

Private Sub txtFld_GotFocus(Index As Integer)
Dim newscroll As Integer
If HScroll1.Max > 0 Then
HScroll1.Value = 0
End If
'If txtFld(Index).Top < 400 And currentrow > -1 And Index < currentrow Or txtFld(Index).Top + txtFld(Index).Height > Frame1.Height Then
'MsgBox txtFld(Index).Top

If txtFld(Index).Top < 0 And Index < currentrow Or txtFld(Index).Top + txtFld(Index).Height > Frame1.Height Then


updatescroll Index
End If
hasentered = False
currentrow = Index



dirty = False
With Screen.ActiveForm
.ActiveControl.SetFocus
.ActiveControl.SelStart = 0
.ActiveControl.SelLength = Len(.ActiveControl.Text)
End With


End Sub

Private Sub txtFld_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'MsgBox colinfo(index)
If KeyCode = 40 And oldtable <> "" And Index + 1 < txtFld.UBound Then
    If check2(Index + 1).Value = 1 Then
        dirty = True
        txtFld(Index + 2).SetFocus
        
        Exit Sub
    End If
End If
If KeyCode = 40 And oldtable <> "" And Index + 1 = txtFld.UBound Then
    If check2(Index + 1).Value = 1 And txtFld(Index + 1).Text <> "" And Command3(Index).caption <> "" Then
    
    
        'dirty = True
        currentrow = currentrow + 1
        processnewrow
        Exit Sub
    End If
End If
If KeyCode = 40 And Index < txtFld.UBound Then
    'processnewrow
    dirty = True
    txtFld(Index + 1).SetFocus
    Exit Sub
End If
'If KeyCode = 40 And Index = txtFld.UBound And colinfo(Index).oldname <> "" Then
  If KeyCode = 40 And Index = txtFld.UBound And txtFld(Index).Text <> "" And Command3(Index).caption <> "" Then
  
    processnewrow
    Exit Sub
End If
'If KeyCode = 38 Then
'MsgBox "test"
'End If

If KeyCode = 38 And oldtable <> "" And Index - 1 > 0 Then
    If check2(Index - 1).Value = 1 Then
        dirty = True
        txtFld(Index - 2).SetFocus
        Exit Sub
    End If
End If
If KeyCode = 38 And oldtable = "" And Index - 1 > 0 Then
    dirty = True
    txtFld(Index - 1).SetFocus
    Exit Sub
    'End If
End If
If KeyCode = 38 And oldtable <> "" And Index > 0 Then
    'MsgBox "test"
    
    dirty = True
    txtFld(Index - 1).SetFocus
    txtFld_GotFocus (Index - 1)
    
    Exit Sub
End If




'If Index > 0 And KeyCode = 38 And oldtable <> "" Or Index > 1 And KeyCode = 38 Then

'set focus up
'txtFld(Index - 1).SetFocus

'dirty = True

'ElseIf KeyCode = 40 And txtFld.UBound > 0 And index < txtFld.UBound Then

'dirty = True
'txtFld(index + 1).SetFocus

'set focus down
'ElseIf KeyCode = 40 And index < txtFld.UBound And colinfo(index).oldname <> "" Then
'ElseIf KeyCode = 40 And Index < txtFld.UBound Then

'dirty = True
'set focus down
'txtFld(Index + 1).SetFocus
'ElseIf KeyCode = 40 And Index = txtFld.UBound And colinfo(Index).oldname <> "" Then

'processnewrow


'End If

End Sub

Private Sub txtFld_LostFocus(Index As Integer)
'MsgBox "test"
If dirty = False Then

For x = 0 To txtFld.UBound

If txtFld(x).Text = txtFld(Index) And x <> Index And txtFld(x).Text <> "" Then

MsgBox "You must enter a unique column name"
txtFld(Index).SetFocus
Exit Sub
Exit For
End If
Next


If IsNumeric(txtFld(Index)) = True Then
MsgBox "Sorry, you must enter the name of the field, not a number"
txtFld(Index).SetFocus
Exit Sub
End If
If txtFld(Index).Text = "" And colinfo(Index).oldname <> "" Then
MsgBox "Sorry, you cannot change the name of the field to blank"
txtFld(Index).SetFocus
Exit Sub
End If

If Command3(Index).caption = "" And Trim(txtFld(Index).Text) <> "" And dirty = False And colinfo(Index).oldname = "" Then



'If Command3(Index).caption = "" And Trim(txtFld(Index).Text) <> "" Then
    Command3_Click (Index)
    ElseIf Index = txtFld.UBound And colinfo(Index).oldname <> "" Then
    processnewrow
    
End If
End If

End Sub

Private Sub VScroll1_Change()
hasentered = False
newx = VScroll1.Height / txtFld(0).Height
searches = InStr(newx, ".")
If searches > 0 Then

newx = Mid(newx, 1, searches - 1)
End If

'If VScroll.Value = 0 Then
'VScroll.LargeChange = newx
'Else

'MsgBox newx
On Error Resume Next

VScroll1.LargeChange = newx - 2
'End If

VScroll1.SmallChange = 2
On Error GoTo 0

'If VScroll.Value <> 0 Then
Dim newposition As Long
If VScroll1.Value < 2 Then
newposition = 0
Else
newposition = VScroll1.Value - 1
newposition = newposition * CLng(-txtFld(0).Height)
End If

'Dim newx As Long

For x = 0 To txtFld.UBound

'MsgBox Text1(x).Text

'Text1(x).Top = (VScroll.Value * -100) + xx
'MsgBox Text1(x)
'MsgBox Text1(x).Top
'On Error GoTo messages
'newx = CLng(VScroll.Value) * CLng(-200)
'newx = newx + CLng(xx)


txtFld(x).Top = newposition
Command3(x).Top = newposition
check2(x).Top = newposition
widths(x).Top = newposition
chkrequired(x).Top = newposition


'Text1(x).Top = newposition
newposition = newposition + txtFld(0).Height

'xx = xx + 300
Next
'MsgBox "test"

If txtFld(txtFld.UBound).Top + txtFld(txtFld.UBound).Height < VScroll1.Top + VScroll1.Height Then
'If txtFld(txtFld.UBound).Top + txtFld(txtFld.UBound).Height > VScroll1.Height + VScroll1.Top - Label3.Height - 300 Then

'MsgBox Text1(100).Top
'MsgBox Form1.Height + StatusBar1.Height
'MsgBox VScroll.Height
'MsgBox VScroll.Top

VScroll1.Max = VScroll1.Max
Else
VScroll1.Max = txtFld.UBound



End If

Exit Sub
End Sub

Private Sub widths_Change(Index As Integer)
hasentered = True

End Sub

Private Sub widths_GotFocus(Index As Integer)
'updatescroll Index
hasentered = False
With Screen.ActiveForm
.ActiveControl.SetFocus
.ActiveControl.SelStart = 0
.ActiveControl.SelLength = Len(.ActiveControl.Text)
End With
dirty = False
Dim newscrolls As Integer
If HScroll1.Max > 0 Then
newscrolls = Command3(0).Left + Command3(0).Width
newscrolls = newscrolls / 100
newscrolls = newscrolls - 2
HScroll1.Value = newscrolls
End If
'currentrow = Index
'MsgBox "test"

If txtFld(Index).Top < 0 And Index < currentrow Or txtFld(Index).Top + txtFld(Index).Height > Frame1.Height Then


updatescroll Index
End If
End Sub

Private Sub widths_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)

'txtFld_KeyDown Index, KeyCode, 0

If KeyCode = 38 Or KeyCode = 40 Then
    If colinfo(Index).oldwidth > widths(Index).Text Then
        MsgBox "Sorry, you must enter a number equal or higher than the previous width"
        widths(Index).SetFocus
        Exit Sub
    End If
    If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
        MsgBox "Sorry, you must enter a valid number"
        widths(Index).SetFocus
        Exit Sub
    End If
    If widths(Index - 1).Visible = True And KeyCode = 38 Then
        dirty = True
        widths(Index - 1).SetFocus
        Exit Sub
    End If
    If KeyCode = 40 And Index < txtFld.UBound Then
    
        If KeyCode = 40 And widths(Index + 1).Visible = True Then
            dirty = True
            widths(Index + 1).SetFocus
            Exit Sub
        End If
    End If
    
txtFld_KeyDown Index, KeyCode, 0
End If













'If Index > 0 And KeyCode = 38 Then
'If colinfo(Index).oldwidth > widths(Index).Text Then
'MsgBox "Sorry, you must enter a number = or higher than the previous width"
'Exit Sub
'End If
'If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
'MsgBox "Sorry, you must enter a valid number"
'widths(Index).SetFocus
'Exit Sub
'End If

'If widths(Index - 1).Visible = True Then

'set focus up
'txtFld(index + 1).SetFocus
'widths(Index - 1).SetFocus

'dirty = True
'Exit Sub
'End If

'ElseIf KeyCode = 40 And Index < txtFld.UBound Then
'If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
'MsgBox "Sorry, you must enter a valid number"
'widths(Index).SetFocus
'Exit Sub
'End If
'If colinfo(Index).oldwidth > widths(Index).Text Then
'MsgBox "Sorry, you must enter a number = or higher than the previous width"
'Exit Sub
'End If

'If widths(Index + 1).Visible = True Then
'set focus down
'widths(Index + 1).SetFocus

'dirty = True
'Exit Sub
'End If






'End If
'If Index = txtFld.UBound And KeyCode = 40 Then
'If IsNumeric(widths(Index).Text) = True Then

'If colinfo(Index).oldwidth > widths(Index).Text Then
'MsgBox "Sorry, you must enter a number = or higher than the previous width"
'Exit Sub
'End If
'End If

'processnewrow
'Exit Sub
'End If


'End If
End Sub
Private Sub detprocess(Index As Integer, widthss As Boolean)
If Index = txtFld.UBound And dirty = False Then

processnewrow
'ElseIf widths(index + 1).Visible = True And widthss = True Then

'widths(index + 1).SetFocus
'ElseIf check2(index + 1).Visible = True And widthss = False Then
'check2(index + 1).SetFocus

ElseIf dirty = False Then

txtFld(Index + 1).SetFocus
End If


End Sub
Private Sub widths_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
MsgBox "Sorry, you must enter a valid number"
widths(Index).SetFocus
Exit Sub
End If
detprocess Index, True

End If

End Sub

Private Sub widths_LostFocus(Index As Integer)




'txtFld_KeyDown Index, KeyCode, 0

'If KeyCode = 38 Or KeyCode = 40 Then
If dirty = False Then

    If colinfo(Index).oldwidth > widths(Index).Text Then
        MsgBox "Sorry, you must enter a number equal or higher than the previous width"
        widths(Index).SetFocus
        Exit Sub
    End If
    If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
        MsgBox "Sorry, you must enter a valid number"
        widths(Index).SetFocus
        Exit Sub
    End If
    'If widths(Index - 1).Visible = True And KeyCode = 38 Then
        'dirty = True
        'widths(Index - 1).SetFocus
        'Exit Sub
    'End If
  If hasentered = False Then
  Exit Sub
  End If
  
    If Index < txtFld.UBound Then
    
        If widths(Index + 1).Visible = True Then
            dirty = True
            widths(Index + 1).SetFocus
            Exit Sub
        End If
    End If
    
    
    
    
txtFld_KeyDown Index, 40, 0
'End If

End If












'If widths(Index).Text <> "" And IsNumeric(widths(Index)) = False Then
'MsgBox "Sorry, you must enter a valid number"
'widths(Index).SetFocus
'Exit Sub
'End If
'detprocess Index, True


'If dirty = False Then
'processnewrow
'End If

End Sub
Private Sub savechanges2()
'MsgBox "test"
Dim tmpCN As New ADODB.Connection   'temporary connection
    Set tmpCN = New ADODB.Connection
    
    Dim scmd As String
    
Dim x As Integer
Dim newColumn As ADOX.Column

    If currentrow = -1 Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If
    
    If txtFld(currentrow).Text <> "" And Command3(currentrow).caption = "" Then
        MsgBox "Sorry, you must choose a column type"
        Load comboss
        Exit Sub
    End If
    For x = 0 To txtFld.UBound
    
        If txtFld(x).Text = txtFld(Index) And x <> Index Then
            MsgBox "You must enter a unique column name"
            txtFld(Index).SetFocus
            Exit Sub
            Exit For
        End If
    Next

    If txtFld(currentrow).Text = "" And colinfo(currentrow).oldname <> "" Then
        MsgBox "You cannot change the name of the column to blank"
        txtFld(currentrow).SetFocus
        Exit Sub
    End If




    'If oldtable = "" Then
        'MsgBox "create new table"
        'Set tblNew = New ADOX.Table
            'tblNew.Name = Text1.Text
         'Set tblNew.ParentCatalog = cat
         'cat.Tables.Append tblNew
         'cat.Tables.Refresh
     
    'End If
    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable <> "" Then
        'MsgBox "will delete table"
        sqls = "drop table " & oldtable
        dbObj.Execute (sqls)
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable = "" Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If oldtable = "" Then
    
    
    
    Dim tblnew As New ADOX.Table
    
    
    Set tblnew = New ADOX.Table
    tblnew.Name = Text1.Text
    
 Set tblnew.ParentCatalog = cat
    x = 0
    'While x <= 9
    
    For x = 0 To txtFld.UBound
    If Command3(x).caption <> "" Then
    
    If IsNumeric(widths(x).Text) = False Then
    widthss = 50
    Else
    widthss = widths(x).Text
    End If
        
        If Command3(x).caption = "Text" Then
            tblnew.Columns.Append txtFld.Item(x).Text, adVarWChar, widthss
            'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'by rizka   - check this later.
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
            
        ElseIf Command3(x).caption = "Integer" And check2(x).Value = 1 Then
        tblnew.Columns.Append txtFld.Item(x).Text, adInteger
        tblnew.Columns(txtFld.Item(x).Text).Properties("AutoIncrement") = True
        'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Integer" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adInteger
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Date" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adDate
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         
        ElseIf Command3(x).caption = "Boolean" Then
        tblnew.Columns.Append txtFld.Item(x).Text, adBoolean
        ElseIf Command3(x).caption = "Currency" Then
        tblnew.Columns.Append txtFld.Item(x).Text, adCurrency
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
        ElseIf Command3(x).caption = "Notes" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adLongVarWChar
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
          'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         'ColMyColum.Attributes = adColNullable
        End If
        
    End If
        'x = x + 1
    'Wend
    'custom tables
    Next
    
    cat.Tables.Append tblnew
    'MsgBox "test"
    
    
    frmMain.SetupDatabase (Text1.Text)
    Unload Me
    Exit Sub
    End If
    




    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable = "" Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    
    
    If oldtable <> Text1.Text And oldtable <> "" Then
        cat.Tables.Item(oldtable).Name = Text1.Text
        cat.Tables.Refresh
    End If
'MsgBox "test"

    For x = 0 To txtFld.UBound
    
    
    'ColName = "tmpColumn"
    
    'colType = oldColType
    'ColWidth = Me.txtNewSize.Text
    'ColWidth = widths(x).Text
    
    'cat.Tables.Item(TblName).Columns.Append ColName, colType, ColWidth
    
    
    
    If Command3(x).caption <> "" And Command3(x).caption <> "Integer" Then
    
    
    
    If IsNumeric(widths(x).Text) = False Then
    widthss = 50
    Else
    widthss = widths(x).Text
    End If
        
        
        If Command3(x).caption = "Text" Then
            cat.Tables(Text1.Text).Columns.Append "tmps", adVarWChar, widthss
            
            'tblnew.Columns.Append txtFld.Item(x).Text, adVarWChar, widthss
            'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'by rizka   - check this later.
            'If Me.chkrequired(x).Value = 0 Then
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = adColNullable
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'Else
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = 0
                
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            'End If
            
            
        ElseIf Command3(x).caption = "Integer" And check2(x).Value = 1 Then
        'tblnew.Columns.Append txtFld.Item(x).Text, adInteger
        cat.Tables(Text1.Text).Columns.Append "tmps", adInteger
        cat.Tables(Text1.Text).Columns("tmps").Properties("AutoIncrement") = True
        
        'tblnew.Columns(txtFld.Item(x).Text).Properties("AutoIncrement") = True
        'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Integer" Then
        cat.Tables(Text1.Text).Columns.Append "tmps", adInteger
         'tblnew.Columns.Append txtFld.Item(x).Text, adInteger
            'If Me.chkrequired(x).Value = 0 Then
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = adColNullable
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'Else
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = 0
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            'End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Date" Then
        cat.Tables(Text1.Text).Columns.Append "tmps", adDate
         'tblnew.Columns.Append txtFld.Item(x).Text, adDate
            'If Me.chkrequired(x).Value = 0 Then
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = adColNullable
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'Else
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = 0
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            'End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         
        ElseIf Command3(x).caption = "Boolean" Then
        cat.Tables(Text1.Text).Columns.Append "tmps", adBoolean
        'tblnew.Columns.Append txtFld.Item(x).Text, adBoolean
        ElseIf Command3(x).caption = "Currency" Then
        cat.Tables(Text1.Text).Columns.Append "tmps", adCurrency
        'tblnew.Columns.Append txtFld.Item(x).Text, adCurrency
            'If Me.chkrequired(x).Value = 0 Then
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = adColNullable
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'Else
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = 0
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            'End If
        ElseIf Command3(x).caption = "Notes" Then
        cat.Tables(Text1.Text).Columns.Append "tmps", adLongVarWChar
         'tblnew.Columns.Append txtFld.Item(x).Text, adLongVarWChar
            'If Me.chkrequired(x).Value = 0 Then
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = adColNullable
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'Else
                'cat.Tables(Text1.Text).Columns("tmps").Attributes = 0
                'tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            'End If
          'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         'ColMyColum.Attributes = adColNullable
        End If
        
        
    
    
    
    'cat.Tables.Item(TblName).Columns.Refresh
    'cat.Tables.Refresh
    
    
    
    
    cat.Tables(Text1.Text).Columns.Refresh
    
    'cat.Tables.Item(TblName).Columns.Refresh
    cat.Tables.Refresh
    If colinfo(x).oldname <> "" Then
    
    scmd = "update " & Text1.Text & " set tmps=" & colinfo(x).oldname
    'MsgBox scmd & "  " & txtFld(x).Text
    
    
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName         'change tmpColumn with a variable - later
    tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    tmpCN.Execute scmd
    tmpCN.Close
    Set tmpCN = Nothing
    cat.Tables.Item(Text1.Text).Columns.Delete colinfo(x).oldname
    cat.Tables.Refresh
    End If
    
    'cat.ActiveConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    
    cat.Tables(Text1.Text).Columns.Item("tmps").Name = txtFld(x).Text
    cat.Tables(Text1.Text).Columns.Refresh
    'End If
    ElseIf Command3(x).caption = "Integer" And colinfo(x).oldname <> txtFld(x).Text Then
    cat.Tables(Text1.Text).Columns(colinfo(x).oldname).Name = txtFld(x).Text
    cat.Tables(Text1.Text).Columns.Refresh
    
    
End If

    'cat.Tables.Item(TblName).Columns.Item("tmpColumn").Name = oldColName
    'cat.Tables.Item(TblName).Columns.Refresh
    'cat.Tables.Refresh
    
    'frmCat.Tables.Item(TblName).Columns.Delete oldColName
    'frmCat.Tables.Refresh
    
    
    
    
    
    Next
    If oldtable <> "" Then
        If UBound(deleterow) > 0 Then
        For x = 1 To UBound(deleterow)
                sqls = "alter table " & oldtable & " drop column " & deleterow(x).oldname
                dbObj.Execute sqls
            Next
        End If
    End If
    Call frmMain.SetupDatabase(Text1.Text)
    Unload Me
    
    Exit Sub

End Sub
Private Sub newsavechanges()
'MsgBox "test"

Dim x As Integer
Dim newColumn As ADOX.Column
Dim tmpCN As New ADODB.Connection   'temporary connection
    Set tmpCN = New ADODB.Connection
    
    Dim scmd As String
    
    
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName         'change tmpColumn with a variable - later
    tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    If currentrow = -1 Then
        MsgBox "No changes saved"
        Call frmMain.SetupDatabase(Text1.Text)
        Unload Me
        
        Exit Sub
    End If
    
    If txtFld(currentrow).Text <> "" And Command3(currentrow).caption = "" Then
        MsgBox "Sorry, you must choose a column type"
        Load comboss
        Exit Sub
    End If
    For x = 0 To txtFld.UBound
    
        If txtFld(x).Text = txtFld(Index) And x <> Index Then
            MsgBox "You must enter a unique column name"
            txtFld(Index).SetFocus
            Exit Sub
            Exit For
        End If
    Next

    If txtFld(currentrow).Text = "" And colinfo(currentrow).oldname <> "" Then
        MsgBox "You cannot change the name of the column to blank"
        txtFld(currentrow).SetFocus
        Exit Sub
    End If




    'If oldtable = "" Then
        'MsgBox "create new table"
        'Set tblNew = New ADOX.Table
            'tblNew.Name = Text1.Text
         'Set tblNew.ParentCatalog = cat
         'cat.Tables.Append tblNew
         'cat.Tables.Refresh
     
    'End If
    'If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable <> "" Then
        'MsgBox "will delete table"
        'sqls = "drop table " & oldtable
        'dbObj.Execute (sqls)
        'tmpCN.Execute (sqls)
        'Call frmMain.SetupDatabase(Text1.Text)
        'Unload Me
        
        'Exit Sub
    'End If
    
    'if txtfld.UBound=0 and check2(0).Value=1 and oldtable<>""
    
    If currentrow = 1 And txtFld.UBound = 1 And txtFld(1).Text = "" And oldtable = "" Then
        MsgBox "No changes saved"
        Call frmMain.SetupDatabase(Text1.Text)
        Unload Me
        
        Exit Sub
    End If

    
    
    
    Dim tblnew As New ADOX.Table
    
    
    Set tblnew = New ADOX.Table
    tblnew.Name = "temps"
 Set tblnew.ParentCatalog = cat
    x = 0
    'While x <= 9
    'If oldtable <> "" Then
        'If UBound(deleterow) > 0 Then
                'For x = 1 To UBound(deleterow)
                'cat.Tables(oldtable).Columns(deleterow(x).oldname).
                'frmCat.Tables.Item(TblName).Columns.Delete oldColName
                'cat.Tables.Item(oldtable).Columns.Delete deleterow(x).oldname
                'cat.Tables.Item(oldtable).Columns.Refresh
                
                'sqls = "alter table " & oldtable & " drop column " & deleterow(x).oldname
                'tmpCN.Execute (sqls)
                'dbObj.Execute sqls
            'Next
        'End If
    'End If
    For x = 0 To txtFld.UBound
    If colinfo(x).oldname <> txtFld(x).Text And colinfo(x).oldname <> "" Then
    
    cat.Tables(oldtable).Columns(colinfo(x).oldname).Name = txtFld(x).Text
    cat.Tables(oldtable).Columns.Refresh
    End If
    
    If Command3(x).caption <> "" Then
    
    If IsNumeric(widths(x).Text) = False Then
    widthss = 50
    Else
    widthss = widths(x).Text
    End If
        
        If Command3(x).caption = "Text" Then
            tblnew.Columns.Append txtFld.Item(x).Text, adVarWChar, widthss
            'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            'by rizka   - check this later.
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
            
        ElseIf Command3(x).caption = "Integer" And check2(x).Value = 1 Then
        tblnew.Columns.Append txtFld.Item(x).Text, adInteger
        tblnew.Columns(txtFld.Item(x).Text).Properties("AutoIncrement") = True
        'tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Integer" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adInteger
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
        ElseIf Command3(x).caption = "Date" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adDate
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
         'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         
        ElseIf Command3(x).caption = "Boolean" Then
        tblnew.Columns.Append txtFld.Item(x).Text, adBoolean
        ElseIf Command3(x).caption = "Currency" Then
        tblnew.Columns.Append txtFld.Item(x).Text, adCurrency
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
        ElseIf Command3(x).caption = "Notes" Then
         tblnew.Columns.Append txtFld.Item(x).Text, adLongVarWChar
            If Me.chkrequired(x).Value = 0 Then
                tblnew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
            Else
                tblnew.Columns(txtFld.Item(x).Text).Attributes = 0
            End If
          'tblNew.Columns(txtFld.Item(x).Text).Attributes = adColNullable
         'ColMyColum.Attributes = adColNullable
        End If
        
        End If
        'x = x + 1
    'Wend
    'custom tables
    Next
    'for x=1 to
    'For x = 1 To UBound(deleterow)
                'MsgBox deleterow(x).oldname
                
                'tblnew.Columns.Append deleterow(x).oldname, adVarWChar, 200
                
                
                'Next
                
                
                
                
                'sqls = "alter table " & Text1.Text & " drop column " & deleterow(x).oldname
                'dbObj.Execute sqls
    
    cat.Tables.Append tblnew
    'MsgBox "test"
    
    If oldtable <> "" Then
    'insert into mytable select * from second table where ID < 1000
    'Dim tmpCN As New ADODB.Connection   'temporary connection
    'Set tmpCN = New ADODB.Connection
    
    'Dim scmd As String
    
    
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName         'change tmpColumn with a variable - later
    'tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
   

    'oldColName = Trim(Me.cmbColumns.Text)
    'oldColType = frmCat.Tables(TblName).Columns(oldColName).Type
    'AddNewColumn
    'Dim oldcolumns As String
    'oldcolumns = ""
    'For x = 0 To txtFld.UBound
    'oldcolumns = oldcolumns & colinfo(x).oldname & ", "
    
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName
    'scmd = "insert into temps select " & colinfo(x).oldname & " from " & oldtable
    'scmd = "insert into temps select " & txtFld(x).Text & " from " & oldtable
    'tmpCN.Execute scmd
    
    'Next
    'oldcolumns = Mid(oldcolumns, 1, Len(oldcolumns) - 2)
    On Error Resume Next
    tmpCN.Close
    Set tmpCN = Nothing
    Set tmpCN = New ADODB.Connection
    
    tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    'scmd = "select * from " & oldtable & " into temps"
    scmd = "insert into temps select * from " & oldtable
    'scmd = "select * from " & oldtable & " into temps"
    
    'scmd = "insert into temps select " & oldcolumns & " from " & oldtable
    'MsgBox scmd
    'sqls = "insert into temps select " & oldcolumns & " from " & oldtable
    'MsgBox sqls
    
    tmpCN.Execute scmd
    'scmd = "UPDATE " & TblName & _
            '" SET tmpColumn = " & oldColName         'change tmpColumn with a variable - later
    'tmpCN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbPath
    'tmpCN.Execute scmd
    
    'If oldtable <> Text1.Text And oldtable <> "" Then
        
    'End If
    
    scmd = "drop table " & oldtable
    tmpCN.Execute scmd
    'tmpCN.Close
    'Set tmpCN = Nothing
    
    
    
    
    End If
    cat.Tables.Item("temps").Name = Text1.Text
    
    
    'If oldtable <> "" Then
        'If UBound(deleterow) > 0 Then
                'For x = 1 To UBound(deleterow)
                'sqls = "alter table " & Text1.Text & " drop column " & deleterow(x).oldname
                'dbObj.Execute sqls
            'Next
        'End If
    'End If
    
    
    
        cat.Tables.Refresh
    'MsgBox "test"
    tmpCN.Close
    Set tmpCN = Nothing
    'Unload Me
    frmMain.SetupDatabase (Text1.Text)
    Unload Me
    
    'Call frmMain.SetupDatabase(Text1.Text)
    
    Exit Sub
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    If oldtable <> Text1.Text And oldtable <> "" Then
        cat.Tables.Item(oldtable).Name = Text1.Text
        cat.Tables.Refresh
    End If

    For x = 0 To txtFld.UBound
        'renaming columns
        If txtFld(x).Text <> "" Then
            If colinfo(x).oldname <> "" And colinfo(x).oldname <> txtFld(x).Text Then
                cat.Tables.Item(Trim(Text1.Text)).Columns.Item(colinfo(x).oldname).Name = txtFld(x).Text
                cat.Tables.Item(Trim(Text1.Text)).Columns.Refresh
                cat.Tables.Refresh
            Else
                ColName = Trim(Me.txtFld(x).Text)
                Select Case LCase(Trim(Command3(x).caption))
                    Case "text"
                        colType = adVarWChar   'adVarChar
                    'Case "float"
                        'colType = adVarNumeric
                    Case "integer"
                        colType = adInteger
                    Case "date"
                        colType = adDate
                    Case "boolean"
                        colType = adBoolean
                    Case "notes"
                        colType = adLongVarWChar
                    Case "currency"
                        colType = adCurrency
                End Select
                 If Command3(x).caption = "Text" And widths(x).Text = "" Then
                     ColWidth = 50
                 Else
                     ColWidth = IIf(Trim(Me.widths(x).Text) = "", 0, widths(x).Text)
                    If colType = adBoolean Then ColWidth = 1
                 End If
                 
                
                Set newColumn = New ADOX.Column
                newColumn.Name = ColName
                newColumn.Type = colType
                newColumn.DefinedSize = ColWidth
                
                If chkrequired(x).Value = 0 Then
                    If colType <> adBoolean Then
                        newColumn.Attributes = adColNullable
                    End If
                End If
                cat.Tables.Item(Text1.Text).Columns.Append newColumn
                If check2(x).Value = 1 Then
                    If colType = adInteger Then
                        'newColumn.Properties("AutoIncrement").Value = True
                        cat.Tables(Text1.Text).Columns(ColName).Properties("autoincrement").Value = True
                    End If
                End If
                cat.Tables.Item(Text1.Text).Columns.Refresh
                cat.Tables.Refresh
                
                Set newColumn = Nothing
                
            End If
        End If
    Next
    
    'MsgBox "test"
    Call frmMain.SetupDatabase(Text1.Text)
    Unload Me
    
    Exit Sub
End Sub

Private Sub MySaveChanges()
'savechanges2
newsavechanges

Exit Sub

Dim x As Integer
Dim newColumn As ADOX.Column

    If currentrow = -1 Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If
    
    If txtFld(currentrow).Text <> "" And Command3(currentrow).caption = "" Then
        MsgBox "Sorry, you must choose a column type"
        Load comboss
        Exit Sub
    End If
    For x = 0 To txtFld.UBound
    
        If txtFld(x).Text = txtFld(Index) And x <> Index Then
            MsgBox "You must enter a unique column name"
            txtFld(Index).SetFocus
            Exit Sub
            Exit For
        End If
    Next

    If txtFld(currentrow).Text = "" And colinfo(currentrow).oldname <> "" Then
        MsgBox "You cannot change the name of the column to blank"
        txtFld(currentrow).SetFocus
        Exit Sub
    End If

    If oldtable = "" Then
        'MsgBox "create new table"
        Set tblnew = New ADOX.Table
            tblnew.Name = Text1.Text
         Set tblnew.ParentCatalog = cat
         cat.Tables.Append tblnew
         cat.Tables.Refresh
     
    End If
    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable <> "" Then
        'MsgBox "will delete table"
        sqls = "drop table " & oldtable
        dbObj.Execute (sqls)
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If currentrow = 0 And txtFld.UBound = 0 And txtFld(0).Text = "" And oldtable = "" Then
        MsgBox "No changes saved"
        Unload Me
        Call frmMain.SetupDatabase(Text1.Text)
        Exit Sub
    End If

    If oldtable <> "" Then
        If UBound(deleterow) > 0 Then
            For x = 1 To UBound(deleterow)
                sqls = "alter table " & oldtable & " drop column " & deleterow(x).oldname
                dbObj.Execute sqls
            Next
        End If
    End If
    If oldtable <> Text1.Text And oldtable <> "" Then
        cat.Tables.Item(oldtable).Name = Text1.Text
        cat.Tables.Refresh
    End If

    For x = 0 To txtFld.UBound
        'renaming columns
        If txtFld(x).Text <> "" Then
            If colinfo(x).oldname <> "" And colinfo(x).oldname <> txtFld(x).Text Then
                cat.Tables.Item(Trim(Text1.Text)).Columns.Item(colinfo(x).oldname).Name = txtFld(x).Text
                cat.Tables.Item(Trim(Text1.Text)).Columns.Refresh
                cat.Tables.Refresh
            Else
                ColName = Trim(Me.txtFld(x).Text)
                Select Case LCase(Trim(Command3(x).caption))
                    Case "text"
                        colType = adVarWChar   'adVarChar
                    'Case "float"
                        'colType = adVarNumeric
                    Case "integer"
                        colType = adInteger
                    Case "date"
                        colType = adDate
                    Case "boolean"
                        colType = adBoolean
                    Case "notes"
                        colType = adLongVarWChar
                    Case "currency"
                        colType = adCurrency
                End Select
                 If Command3(x).caption = "Text" And widths(x).Text = "" Then
                     ColWidth = 50
                 Else
                     ColWidth = IIf(Trim(Me.widths(x).Text) = "", 0, widths(x).Text)
                    If colType = adBoolean Then ColWidth = 1
                 End If
                 
                
                Set newColumn = New ADOX.Column
                newColumn.Name = ColName
                newColumn.Type = colType
                newColumn.DefinedSize = ColWidth
                
                If chkrequired(x).Value = 0 Then
                    If colType <> adBoolean Then
                        newColumn.Attributes = adColNullable
                    End If
                End If
                cat.Tables.Item(Text1.Text).Columns.Append newColumn
                If check2(x).Value = 1 Then
                    If colType = adInteger Then
                        'newColumn.Properties("AutoIncrement").Value = True
                        cat.Tables(Text1.Text).Columns(ColName).Properties("autoincrement").Value = True
                    End If
                End If
                cat.Tables.Item(Text1.Text).Columns.Refresh
                cat.Tables.Refresh
                
                Set newColumn = Nothing
                
            End If
        End If
    Next
    Unload Me
    Call frmMain.SetupDatabase(Text1.Text)
    Exit Sub
   
End Sub
