VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PARAMETROI 
   BackColor       =   &H00FF0000&
   Caption         =   "Form2"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   510
   ClientWidth     =   14190
   LinkTopic       =   "Form2"
   ScaleHeight     =   6585
   ScaleWidth      =   14190
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "≈ÓÔ‰ÔÚ"
      Height          =   450
      Left            =   11880
      TabIndex        =   0
      Top             =   5940
      Width           =   1425
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   348
      Left            =   3960
      Top             =   6120
      Visible         =   0   'False
      Width           =   2796
      _ExtentX        =   4921
      _ExtentY        =   609
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5760
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13212
      _ExtentX        =   23310
      _ExtentY        =   10160
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   33023
      HeadLines       =   1
      RowHeight       =   22
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   161
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "–¡—¡Ã≈‘—œ…"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "VAR"
         Caption         =   " Ÿƒ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "SXOLIA"
         Caption         =   "”˜¸ÎÈ·"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "TIMH"
         Caption         =   "TÈÏﬁ"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "KATEG"
         Caption         =   " ¡‘«√œ—…¡"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1032
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         Size            =   349
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
      EndProperty
   End
   Begin VB.Label PARAM 
      Height          =   276
      Left            =   180
      TabIndex        =   2
      Top             =   5940
      Visible         =   0   'False
      Width           =   3096
   End
End
Attribute VB_Name = "PARAMETROI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   PARAMETROI.PARAM.Caption = "APOT2"
'   PARAMETROI.SHOW 1
'    F_TAB = Val(FindParametroi(1,"APOT2", "F_TAB", "3", "”Â ÔÈ¸ TAB ÂﬂÌ·È ÛÙ·Ï·ÙÁÏ›ÌÔ"))
'Function FindParametroi(FORMA As String, parametros As String, default As String, sxolia As String)
'   Dim R As New ADODB.Recordset
'   Dim sql As String
'
'   If default = "DELETE" Then
'     Gdb.Execute "DELETE FROM PARAMETROI WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
'     FindParametroi = 0
'     Exit Function
'   End If
'   sql = "select * from PARAMETROI where FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
'   R.Open sql, Gdb, adOpenDynamic, adLockBatchOptimistic
'
'   If R.EOF Then
'      sql = "insert into PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('" + FORMA + "','" + parametros + "','" + default + "','" + sxolia + "')"
'      Gdb.Execute sql
'
'      FindParametroi = default
'   Else
'
'      FindParametroi = R("TIMH")
'
'
'   End If
'
'
'
'
'
'End Function

Dim F_DOK

Private Sub Command1_Click()

        '<EhHeader>
        On Error GoTo Command1_Click_Err

        '</EhHeader>

100     Unload Me

        '<EhFooter>
        Exit Sub

Command1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Command1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       ' SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Command1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub DataGrid1_Click()

        ' DATAGRID1.SelEndCol=0
        '<EhHeader>
        On Error GoTo DataGrid1_Click_Err

        '</EhHeader>

100     DataGrid1.Col = 2
110     DataGrid1.SelStart = 0
120     DataGrid1.SelLength = Len(DataGrid1.Text)

        '    TIMText3.SelStart = 0
        '    TIMText3.SelLength = Len(TIMText3.Text)

        '<EhFooter>
        Exit Sub

DataGrid1_Click_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.DataGrid1_Click " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
       ' SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.DataGrid1_Click " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Load()

        '<EhHeader>
        On Error GoTo Form_Load_Err

        '</EhHeader>

100     Adodc1.ConnectionString = gCONNECT

        ' DataGrid1.C= False
        'DataGrid1.Splits(1).Columns(0).Visible = False

'         If Len(PARAM.Caption) > 0 Then
'            Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "'"
'         Else
'            Adodc1.RecordSource = "select * FROM PARAMETROI"
'         End If
'            Adodc1.Refresh
'        '    Adodc1.Recordset.MoveFirst
        '<EhFooter>
        Exit Sub

Form_Load_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Form_Load " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
      '  SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Form_Load " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub

Private Sub Form_Paint()

        '<EhHeader>
        On Error GoTo Form_Paint_Err

        '</EhHeader>
        On Error Resume Next

100     If Len(PARAM.Caption) > 0 Then
110         Adodc1.RecordSource = "select * FROM PARAMETROI WHERE FORMA='" + PARAM.Caption + "' ORDER BY KATEG,SXOLIA"
        Else
120         Adodc1.RecordSource = "select * FROM PARAMETROI"
        End If

130     Adodc1.Refresh
140     Adodc1.Recordset.MoveFirst
150     DataGrid1.Columns(1).Width = 7000

        '<EhFooter>
        Exit Sub

Form_Paint_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.PARAMETROI.Form_Paint " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
      '  SAVE_ERROR Err.Description & " in ADOMERCNEW.PARAMETROI.Form_Paint " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Sub
