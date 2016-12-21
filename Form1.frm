VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form form1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Form1"
   ClientHeight    =   8190
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   11760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   11760
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9600
      Top             =   5280
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\User\Desktop\LAG_EURO.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\User\Desktop\LAG_EURO.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "RTData"
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
   Begin VB.CommandButton Command2 
      Caption         =   "SyloghKinPos"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7320
      TabIndex        =   13
      Top             =   4080
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   3840
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   661
      _Version        =   393216
      Format          =   52232193
      CurrentDate     =   42689
   End
   Begin VB.CheckBox specialDay 
      BackColor       =   &H00C0FFC0&
      Caption         =   "≈ÎÂ„˜ÔÚ Í·È ÙÁÚ ÁÏ›Ò·Ú"
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   3360
      Width           =   2655
   End
   Begin VB.CommandButton PONTOI 
      Caption         =   "–œÕ‘œ…"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7200
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CheckBox EID2 
      Caption         =   "≈Õ«Ã≈—Ÿ”« ≈…ƒŸÕ"
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton EID 
      Caption         =   "≈Õ«Ã≈—Ÿ”« œÀŸÕ ‘ŸÕ ≈…ƒŸÕ"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7080
      TabIndex        =   5
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox Check 
      Caption         =   "≈ÌÂÒ„ÔÔÈÁÛÁ ≈Ó¸‰Ôı"
      Height          =   255
      Left            =   7080
      TabIndex        =   4
      Top             =   6360
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Height          =   6690
      Left            =   360
      TabIndex        =   2
      Top             =   360
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Interval        =   20000
      Left            =   1320
      Top             =   3720
   End
   Begin VB.CommandButton Command1 
      Caption         =   "≈Œœƒœ”"
      Height          =   495
      Left            =   7080
      TabIndex        =   0
      ToolTipText     =   "¡Ò˜ÂﬂÔ ÏÂ servers :    C:\MERCVB\REMOTE.TXT"
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label f_apotL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   5520
      TabIndex        =   19
      Top             =   120
      Width           =   45
   End
   Begin VB.Label lblBARCODES¡–œ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C000C0&
      Caption         =   "BARCODES ¡–œ  ≈Õ‘—… œ"
      Height          =   195
      Left            =   10080
      TabIndex        =   18
      Top             =   1440
      Width           =   1995
   End
   Begin VB.Label lbl≈…ƒ«¡–œ 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "≈…ƒ« ¡–œ  ≈Õ‘—… œ"
      Height          =   195
      Left            =   10080
      TabIndex        =   17
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Label lblKINHSEISSE 
      AutoSize        =   -1  'True
      BackColor       =   &H00FF0000&
      Caption         =   "KINHSEIS SE KENTRIKO"
      Height          =   195
      Left            =   10080
      TabIndex        =   16
      Top             =   720
      Width           =   1875
   End
   Begin VB.Label lblENHMERVSH¡–œ 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "ENHMERVSH ¡–œ POS"
      Height          =   195
      Left            =   10080
      TabIndex        =   15
      Top             =   360
      Width           =   1770
   End
   Begin VB.Label lblLabel4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      Height          =   195
      Left            =   7200
      TabIndex        =   14
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label lTameio 
      BackColor       =   &H0080FFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "◊Ÿ—…” ≈Õ«Ã≈—Ÿ”« ≈…ƒŸÕ //‘œ DELEEGGTIM NA SBHNEI MIA FORA TO XRONO"
      Height          =   495
      Left            =   3270
      TabIndex        =   9
      Top             =   7485
      Width           =   6660
   End
   Begin VB.Label Label2 
      Caption         =   "v.130125 ME DELEGGTIM"
      Height          =   255
      Left            =   8880
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "ver.2016.09"
      Height          =   252
      Index           =   1
      Left            =   624
      TabIndex        =   3
      Top             =   7440
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "REMOTE.TXT DSN  ≈Õ‘—… OY H/Y"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

      #If Win32 Then

        Private Declare Function SQLAllocEnv Lib "odbc32.dll" (phenv As Long) _
            As Integer
        Private Declare Function SQLAllocConnect Lib "odbc32.dll" (ByVal henv _
            As Long, phdbc As Long) As Integer
        Private Declare Function SQLConnect Lib "odbc32.dll" (ByVal hdbc As _
            Long, ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As _
            String, ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal _
            cbAuthStr As Integer) As Integer
        Private Declare Function SQLFreeEnv Lib "odbc32.dll" (ByVal henv As _
            Long) As Integer
        Private Declare Function SQLFreeConnect Lib "odbc32.dll" (ByVal hdbc As _
            Long) As Integer
        Private Declare Function SQLError Lib "odbc32.dll" (ByVal henv As Long, _
            ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As String, _
            pfNativeError As Long, ByVal szErrorMsg As String, ByVal cbErrorMsgMax _
            As Integer, pcbErrorMsg As Integer) As Integer
      #ElseIf Win16 Then
        Private Declare Function SQLAllocEnv Lib "odbc.dll" (phenv As Long) As _
            Integer
        Private Declare Function SQLAllocConnect Lib "odbc.dll" (ByVal henv As _
            Long, phdbc As Long) As Integer
        Private Declare Function SQLConnect Lib "odbc.dll" (ByVal hdbc As Long, _
            ByVal szDSN As String, ByVal cbDSN As Integer, ByVal szUID As String, _
            ByVal cbUID As Integer, ByVal szAuthStr As String, ByVal cbAuthStr As _
            Integer) As Integer
        Private Declare Function SQLFreeEnv Lib "odbc.dll" (ByVal henv As Long) _
            As Integer
        Private Declare Function SQLFreeConnect Lib "odbc.dll" (ByVal hdbc As _
            Long) As Integer
        Private Declare Function SQLError Lib "odbc.dll" (ByVal henv As Long, _
            ByVal hdbc As Long, ByVal hstmt As Long, ByVal szSqlState As String, _
            pfNativeError As Long, ByVal szErrorMsg As String, ByVal cbErrorMsgMax _
            As Integer, pcbErrorMsg As Integer) As Integer
      #End If

      Private Const SQL_SUCCESS As Long = 0
      Private Const SQL_SUCCESS_WITH_INFO As Long = 1

Dim f_ExistOnServer As Integer
Dim F_SQL As String

Dim GDIR

Dim CC As String


Dim F_LOADPEL As Integer

Dim F_MHXANH As String
Dim F_APOT As String
Dim F_POS As Integer  'ENHME—Ÿ”« ¡–œ POS
Dim F_PONTOI As Integer  'ENHME—Ÿ”« –ONTŸÕ
Dim F_ENHM_EIDH As Integer  ' AR.APOTHIKIS

Dim F_SYNCHRO_KIN As Integer ' = 1


Dim SERV1 As Integer
Dim SERV2 As Integer
Dim SERV3 As Integer

Dim FTAM_FIELD As String ' T30,T40,T41...

 Dim fRL As New ADODB.Recordset

Dim F_RUNNING ' FLAG GIA NA MHN TRTEXEI DYO FORES H ENHMERVSH

Private Sub synchro()
    '≈Õ«Ã≈—Ÿ”«  …Õ«”≈ŸÕ
    '–—œ”œ◊« ≈◊Ÿ ¬¡À≈… ”‘œ REMOTE EGGTIMDEMO ANTI EGGTIM (SEIRES 50 & 130 )
    'META TIS DOKIMES NA GINEI EGGTIM  12/8/2012
    ' H HMEROMHNIA EINAI '12/16/2012'
    List1.BackColor = lblKINHSEISSE.BackColor

    Dim K

    Dim fname

    Dim z

10  On Error GoTo SEELINE

20  F_RUNNING = 1

30  GDB.Execute "UPDATE EGGTIM SET KOLA=" + F_APOT + " WHERE HME>= '12/16/2016' ;"

40  GDB.Execute "UPDATE EGGTIM SET APOT=" + F_APOT + ",XRE=0,PIS=POSO  WHERE HME>= '12/16/2016' AND LEFT(ATIM,1)='L' ;"

    Dim R  As New ADODB.Recordset

    Dim RR As New ADODB.Recordset

    Dim CC

    ' ‚ÒÈÛÍ˘ ÙÔ ÏÂ„·Î˝ÙÂÒÔ ID –œ’ ’–¡—◊≈… ”‘«Õ ƒ—¡Ã¡ ¡–œ ’–œ  4 (KOLA=4)
50  RR.Open "SELECT MAX(ID)  FROM EGGTIM WHERE KOLA=" + F_APOT + " AND HME>='12/16/2016' ", GDBR, adOpenDynamic, adLockOptimistic

60  If IsNull(RR(0)) Then
70      CC = 0
80  Else
90      CC = RR(0)
100     End If

110     RR.Close

120     List1.AddItem " MAX(ID) " + Str(CC)

130     R.Open "SELECT count(*) FROM EGGTIM WHERE ID>" + Str(CC) + " AND HME>='12/16/2016'", GDB, adOpenDynamic, adLockOptimistic
140     K = R(0)
150     R.Close

160     List1.AddItem " »¡ ≈Õ«Ã≈—Ÿ»œ’Õ " + Str(K) + " ≈√√—¡÷≈”"

        Dim r5      As New ADODB.Recordset

        Dim R5COUNT As Long

        Dim N       As Long
        
        List1.AddItem Format(Now, "HHmm") + "======="

        'ÏÈ· ˆÔÒ· ÙÁÌ ÁÏÂÒ· ÂÎÂ„˜ÂÈ ÔÎÂÚ ÙÈÚ Â„„Ò·ˆ›Ú
        If Format(Now, "HHmm") = "2300" Then
            List1.AddItem Format(Now, "HHmm") + "/////======="
            Me.Caption = "ÂÎÂ„˜ÔÚ ÙÁÚ " + Format(Now, "DD/MM/YYYY")
             
            If Len(Dir(App.Path + "\ALLYEAR.TXT", vbNormal)) > 0 Then
                R.Open "SELECT * FROM EGGTIM WHERE YEAR(HME)=" + Format(Now, "YYYY") + " ORDER BY HME ", GDB, adOpenDynamic, adLockOptimistic
            Else
                R.Open "SELECT * FROM EGGTIM WHERE HME='" + Format(Now, "MM/DD/YYYY") + "'", GDB, adOpenDynamic, adLockOptimistic
            End If

            'Í·ËıÛÙÂÒÁÛÁ „È·  Ì· ÙÔ ÙÒÂ˜ÂÈ ÏÈ· ˆÔÒ· ÙÁÌ ÁÏÂÒ· KAI OXI >=2
            For N = 1 To 60
                MILSEC 1000
            Next
             
        ElseIf specialDay.Value = vbChecked Then
            specialDay.Value = vbUnchecked
            R.Open "SELECT * FROM EGGTIM WHERE HME='" + Format(DTPicker1.Value, "MM/DD/YYYY") + "'", GDB, adOpenDynamic, adLockOptimistic
        Else

170         R.Open "SELECT * FROM EGGTIM WHERE ID>" + Str(CC) + " AND HME>='12/16/2016'", GDB, adOpenDynamic, adLockOptimistic
        End If

180     RR.Open "SELECT TOP 10 * FROM EGGTIM ", GDBR, adOpenDynamic, adLockOptimistic

190     Do While Not R.EOF
200         DoEvents

            'ÿ¡◊ÕŸ Õ¡ ƒŸ  Ã«–Ÿ” ’–¡—◊≈… «ƒ« ‘œ ID STHN DRAMA
            r5.Open "select count(*) from EGGTIM WHERE ID=" + Str(R("ID")), GDBR, adOpenDynamic, adLockOptimistic
            R5COUNT = r5(0)
            r5.Close
        
            If R5COUNT = 0 Then
210             RR.AddNew

220             For K = 0 To R.Fields.Count - 1
230                 fname = R.Fields(K).Name    ' p.x. FNAME=epo    R(0).NAME

240                 If IsNull(R(K)) Then
250                 Else

260                     If fname = "ATIM2" Then  ' ƒ≈Õ ’–¡—◊≈… ¡Õ‘…”‘œ…◊œ –≈ƒ…œ ”‘«Õ ƒ—¡Ã¡
                        Else
                            RR(fname) = R(fname)    ' rsqk("epo")=r(0)
                        End If
270                 End If

280             Next

290             RR.Update

            End If

            Me.Caption = "ÂÎÂ„˜ÔÚ ÙÁÚ " + Format(R!HME, "DD/MM/YYYY")
300         z = z + 1

310         If z Mod 10 = 0 Then
320             Me.Caption = z
330         End If

340         R.MoveNext

350         DoEvents

360     Loop

370     RR.Close
380     R.Close

390     If List1.ListCount > 30 Then List1.Clear

400     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + " ≈Õ«Ã≈—Ÿ»« ¡Õ " + Str(z) + " ≈√√—¡÷≈”"

410     F_RUNNING = 0

420     Exit Sub

SEELINE:
        'HandleError "MdiForm-load"
        'Resume Next
430     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"

440     On Error Resume Next

450     GDB.Close
460     GDBR.Close

SAVE_ERROR Err.Description & " in SYNCHRO_Click " & "at line " & Erl


470     Me.Caption = Str(Erl) + "-----" + Err.Description
480     open_data

490     Exit Sub    'Resume Next

End Sub

Private Sub Command1_Click()
   g_stop = 1

500     Unload Me
        
End Sub

Private Sub Command2_Click()


   Get_Kin "3"
   
   Get_Kin "2"
   
   Get_Kin "1"
   

   
   
   
End Sub

Sub Get_Kin(ByVal pos)
Dim APOdTEL As String, DATETEL As String  ' ‘≈À≈’‘¡…¡ ¡–œƒ≈…Œ«

On Error GoTo SEELINE

    Dim gMDB As New ADODB.Connection

    gMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\POS" + pos + "\TEC_POS\DATA\LAG_EURO.mdb;Persist Security Info=False"

   'gMDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\DESKTOP-LHRR90N\tec_pos\Data\LAG_EURO.mdb;Persist Security Info=False"

    List1.BackColor = lblENHMERVSH¡–œ.BackColor
    
    
    
    
    Dim R As New ADODB.Recordset

    ' r.Open "select * from RTData where ActionFlag like '0%' order by ID2 ", gMDB, adOpenDynamic, adLockOptimistic



    'ÓÂÍÈÌ‹˘ ·¸ ÙÔ Ù›ÎÔÚ ÏÁ˘Ú ›˜ÂÈ ·¸‰ÂÈÓÁ ÛÙÔÌ ·›Ò· Ôı ‰ÂÌ ÂÍÎÂÈÛÂ
    R.Open "select * from RTData where ActionFlag  order by ID2 DESC ", gMDB, adOpenDynamic, adLockOptimistic
    
    If Not R.EOF Then
        APOdTEL = R!receipt
        DATETEL = R!Date
    End If


     Do While Not R.EOF
     
     
     
     
         If R("ACTIONCODE") = "699" Then
              Exit Do
         End If
         R.MoveNext
     Loop
     
 


   Dim MAXID2 As Long
   
   
   'Ã≈◊—… ¡’‘œ ‘œ ID »¡  …Õ«»≈…”
   MAXID2 = R!ID2
   
   
   
   
   
  R.Close
  
  R.Open "select * from RTData where ActionFlag like '0%' AND ID2<=" + Str(MAXID2) + "  order by ID2 ", gMDB, adOpenDynamic, adLockOptimistic





    Dim apod  As String

    Dim mHME  As String

    Dim r0    As New ADODB.Recordset

    Dim sql   As String

    Dim cHME  As String

    Dim NAFF  As Integer

    Dim OK    As Integer

    Dim cFPA  As String

    Dim CC    As String

    Dim GRAFO As Integer
    
    
    Dim MEKPT As String
    Dim MAXIAKI As Integer
    
    
    
    Do While Not R.EOF
        apod = R!receipt
        mHME = R!Date
        'r0.Open "SELECT * FROM RTDATA WHERE ACTIONCODE='699' AND DATE='" + mHME + "' AND RECEIPT='" + apod + "'", gMDB, adOpenDynamic, adLockOptimistic

'        '≈◊≈… ”’ÕœÀœ ¡—¡ √—¡÷Ÿ ”‘œ EGGTIM
'        If r0.RecordCount > 0 Then GRAFO = 1 Else GRAFO = 0  ' ƒ≈Õ ≈◊≈… ”’ÕœÀ¡ (œÀ… « ¡ ’—Ÿ”«)
       ' r0.Close
        
       'r0.Close
       r0.Open "SELECT count(*) FROM RTDATA WHERE ACTIONCODE='699' AND DATE='" + mHME + "' AND RECEIPT='" + apod + "'", gMDB, adOpenDynamic, adLockOptimistic

        '≈◊≈… ”’ÕœÀœ ¡—¡ √—¡÷Ÿ ”‘œ EGGTIM
        If r0(0) > 0 Then
           GRAFO = 1
        Else
           GRAFO = 0  ' ƒ≈Õ ≈◊≈… ”’ÕœÀ¡ (œÀ… « ¡ ’—Ÿ”«)
           
           '¡Õ « ¡–œƒ≈…Œ« ≈…Õ¡… Ã…”œ‘≈À≈…ŸÃ≈Õ« ¬√≈”  ¡… ¡”≈ ‘«Õ Õ¡ ‘≈À≈…Ÿ”≈…
           If apod = APOdTEL And mHME = DATETEL Then
               Exit Do
           End If
           
           
           
        End If
        r0.Close
        
        
        
        
        
        
        
        OK = 1
        
        'LOOP ”‘«Õ …ƒ…¡ ¡–œƒ≈…Œ«
        Do While apod = R!receipt And mHME = R!Date And Not R.EOF
            
            
           On Error GoTo 0
            
            
            '≈À≈√◊œ” Ã«–Ÿ” ¡ œÀœ’»≈… ≈ –‘Ÿ”«
            MEKPT = "0"
            R.MoveNext
            
            If R.EOF Then
              R.MoveLast
            Else
              If R!ACTIONCODE = "201" Then
                MEKPT = R!ACTIONDATA3
                If R!ACTIONDATA2 = "3" Then
                    'MAXIAKI = 1
                    MEKPT = Replace(Str(Val(MEKPT) / Val(Replace(Left(CNULL(R!ACTIONDATA6), 10), ",", "."))), ",", ".")
                Else
                    'MAXIAKI = 0
                End If
              End If
              R.MovePrevious
            End If
            
            'If GRAFO = 1 Then
            If R!ACTIONCODE = "100" Or R!ACTIONCODE = "110" Or R!ACTIONCODE = "111" Or R!ACTIONCODE = "112" Then
                cHME = IIf(IsNull(R!Date), Format(Now, "DD/MM/YYYY"), R!Date)
                cHME = Mid$(cHME, 4, 2) + "/" + Mid$(cHME, 1, 2) + "/20" + Mid$(cHME, 7, 2)
              
                CC = Left(CNULL(R!ACTIONDATA3) + " ", 1)

                If CC = "1" Then
                    cFPA = "4"
                ElseIf CC = "2" Then
                    cFPA = "3"
                ElseIf CC = "3" Then
                    cFPA = "2"
                Else
                    cFPA = "2"
                End If
                
                
              
                    
                sql = "INSERT INTO EGGTIM (ATIM,KODE,POSO,TIMM,EIDOS,HME,MIK_AJIA,FPA,APOT,EKPT,ATIM2) VALUES ("
                sql = sql + "'L" + apod + "'," 'ATIM
                 sql = sql + "'" + CNULL(R!ACTIONDATA1) + "'," 'KODE
                sql = sql + "" + Replace(CNULL(R!ACTIONDATA5), ",", ".") + "," ' POSO
                sql = sql + "" + Replace(Left(CNULL(R!ACTIONDATA6), 10), ",", ".") + "," 'TIMM
                sql = sql + "'e',"
                sql = sql + "'" + cHME + "',"  ' HME
                sql = sql + "" + Replace(Left(CNULL(R!ACTIONDATA6), 10), ",", ".") + "," 'MIK_AJIA
                sql = sql + "" + cFPA + ","  '÷–¡
                sql = sql + F_APOT + ","  ' ¡–œ»« «
                sql = sql + MEKPT + ","  ' ≈ –‘Ÿ”«
                sql = sql + "'" + pos + "')"
                    
                If GRAFO = 1 Then ' kanonikh apodeixi Ô˜È ÔÎÈÍÁ ·ÍıÒ˘ÛÁ
                    GDB.Execute sql, NAFF

                    If NAFF = 0 Then
                        OK = 0
                    End If
                End If
            End If

            'End If
            
            If R!ACTIONCODE = "411" Then   '–œÕ‘œ…
                cHME = IIf(IsNull(R!Date), Format(Now, "DD/MM/YYYY"), R!Date)
                cHME = Mid$(cHME, 4, 2) + "/" + Mid$(cHME, 1, 2) + "/20" + Mid$(cHME, 7, 2)
                    
                cFPA = "5"
                sql = "INSERT INTO EGGTIM (ATIM,KODE,POSO,TIMM,EIDOS,HME,MIK_AJIA,FPA,APOT,ATIM2) VALUES ("
                sql = sql + "'L" + apod + "'," 'ATIM
                sql = sql + "'" + CNULL(R!ACTIONDATA1) + "'," 'KODE
                sql = sql + "0," ' POSO
                sql = sql + "0," 'TIMM
                sql = sql + "'e',"
                sql = sql + "'" + cHME + "',"  ' HME
                sql = sql + "0," 'MIK_AJIA
                sql = sql + "" + cFPA + ","  '÷–¡
                sql = sql + F_APOT + ","  ' ¡–œ»« «
                sql = sql + "'" + pos + "')"
                    
                If GRAFO = 1 Then ' kanonikh apodeixi Ô˜È ÔÎÈÍÁ ·ÍıÒ˘ÛÁ
                    GDB.Execute sql, NAFF

                    If NAFF = 0 Then
                        OK = 0
                    End If

                    
                    sql = "INSERT INTO DIATAKT ( ARIUMOS,AJIA,SEIRA,PELATHS,HME,ATIM,PONTOI) values ("
                    sql = sql + apod + ","   'ATIM + ","
                    sql = sql + Replace(Str(R!ACTIONDATA5), ",", ".") + ","
                    sql = sql + "'" + pos + "',"
                    sql = sql + "'" + CNULL(R!ACTIONDATA1) + "',"   ' PELATHS
                    sql = sql + "'" + cHME + "',"  ' HME
                    sql = sql + "'L" + apod + "'," 'ATIM
                    sql = sql + CNULL(R!ACTIONDATA3) + ")"
                    GDB.Execute sql
                       
                End If
            End If

            R.MoveNext
            If R.EOF Then Exit Do
        Loop
        
        gMDB.Execute "UPDATE RTDATA SET ACTIONFLAG='*'+MID(ACTIONFLAG,2,1) WHERE DATE='" + mHME + "' AND RECEIPT='" + apod + "'"
  
      ' On Error Resume Next
        If R.EOF Then Exit Do
        
        List1.AddItem R!ACTIONCODE + R!actiondesc

       ' r.MoveNext
    Loop


Exit Sub

SEELINE:



        List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"

        '      Resume Next
        On Error Resume Next

        GDB.Close
        GDBR.Close
        Me.Caption = Str(Erl) + "-----" + Err.Description
        SAVE_ERROR Err.Description & " in GET_KIN " & "at line " & Erl
        open_data

        Exit Sub    'Resume Next





End Sub



'Sub  FileCopy "C:\MERCVB\customer.upd", "\\POS1\TEC_POS\DATA\customer.upd"
'                     FileCopy "C:\MERCVB\customer.upd", "\\POS2\TEC_POS\DATA\customer.upd"
'                     FileCopy "C:\MERCVB\customer.upd", "\\POS3\TEC_POS\DATA\customer.upd"
'
'                     FileCopy "C:\MERCVB\points.upd", "\\POS1\TEC_POS\DATA\points.upd"
'                      FileCopy "C:\MERCVB\points.upd", "\\POS2\TEC_POS\data\points.upd"
'                     FileCopy "C:\MERCVB\points.upd", "\\POS3\TEC_POS\data\points.upd"()
'        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then
'
''OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
''OK----01       036         30                                    –≈—…√—¡÷«                                                              G
''OK----03       087         08                                    ‘…Ã«                                                                D,###.00
''OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
''OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
''OK----17       095         01                                    ‘Ã«Ã¡                                                              I
''----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
''----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
''----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
''OK  ----25       00115                                    ¬  Ÿƒ… œ”
'
'
'
'
'Dim R22 As New ADODB.Recordset
'
'
'            Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2
'
'            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
'             ' ena eidos
'              R22.Open "SELECT  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%' ) ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
'
'             ' ola ta eidh
'             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
'                'On Error GoTo iliadis_error
'
'
'Dim mLTI5 As Single
'Dim MONO As String
'Dim mkod As String
'Dim mbar As String
'Dim MFPA As Integer
'Dim MBARC As String
'
'
'Dim nc As Long
'
'
'            Do While Not R22.EOF
'                ' Open "C:\POSFILE.TXT" For Output As #1
'
'             If IsNull(R22(0)) Then
'                mkod = "."
'             Else
'                mkod = R22(0)
'             End If
'
'
'              If mkod = "005.429" Then
'                 nc = nc + 1
'              End If
'
'             mkod = Replace(mkod, Chr(10), " ")
'             mkod = Replace(mkod, Chr(13), " ")
'
'
'             CC = Left(mkod + Space(21), 20)  ' KVDIKOS
'
'
'
'             'R22 ("ONO")
'             If IsNull(R22("ono")) Then
'                MONO = "."
'             Else
'                MONO = R22("ono")
'             End If
'             MONO = Replace(MONO, Chr(10), " ")
'             MONO = Replace(MONO, Chr(13), " ")
'
'             'R22 ("ONO")
'             If IsNull(R22(1)) Then
'                MBARC = "0000"
'             Else
'                MBARC = R22(1)
'             End If
'
'             MBARC = Replace(MBARC, Chr(10), " ")
'             MBARC = Replace(MBARC, Chr(13), " ")
'
'
'
'             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
'
'
'
'
'
'             CC = CC + Mid(MONO + String(30, " "), 1, 30) + String(21, " ") ' Space(21)
'             If IsNull(R22("LTI5")) Then
'                mLTI5 = 0
'             Else
'                mLTI5 = R22("LTI5")
'             End If
'
'             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
'
'                '  If IsNull(R22("mon")) Then
'                '     CC = CC + Space(3)
'                '   Else
'                '      CC = CC + R22("mon") '+ Space(3), 3)
'                ' End If
'
'
'
'             If IsNull(R22("FPA1")) Then
'                MFPA = 3
'             Else
'                MFPA = R22("FPA1")
'             End If
'
'
'             DoEvents
'
'             Me.Caption = nc
'             nc = nc + 1
'
'             CC = CC + Format(MFPA, "0") + Space(10)
'             CC = CC + " 1"    ' monada Format(R22("mon")
'
'
'   '          On Error Resume Next
'
'            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
'
'             CC = CC + "                       00    0  0      "   ' ÔÌÙÔÈ Êı„ÈÊÔÏÂÌ·
'             CC = Replace(CC, Chr(13), "")
'             If Len(CC) < 140 Then
'                 Print #2, CC
'             Else
'                 Print #1, CC
'             End If
''123456789012340078900230567890
'
'                R22.MoveNext
'            Loop
'
'
'
'
'
'
'             Close #1
'
'             Close #2
'
'
'
'             R22.Close
'
'
'             On Error Resume Next
'             Dim ANS As Integer
'             ANS = MsgBox("Ì· ·ÔÛÙ·ÎÔ˝Ì ÛÙ· pos? ", vbYesNo)
'             If ANS = vbYes Then
'                 If Len(Trim(F_POS1_FOLDER)) > 1 Then
'                     FileCopy "C:\MERCVB\POSFILE.TXT", F_POS1_FOLDER + "\POSFILE.TXT"
'                 Else
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS3\TEC_POS\DATA\POSFILE.TXT"
'                 End If
'             End If
'
'
'' LOAD_PELATES
'
'
'
'
'                On Error Resume Next
'
'                Exit Sub
'
'         '   End If
'
'End Sub
'Sub OLD3LOAD_PELATES()
'        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then
'
''OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
''OK----01       036         30                                    –≈—…√—¡÷«                                                              G
''OK----03       087         08                                    ‘…Ã«                                                                D,###.00
''OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
''OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
''OK----17       095         01                                    ‘Ã«Ã¡                                                              I
''----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
''----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
''----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
''OK  ----25       00115                                    ¬  Ÿƒ… œ”
'
'
'
'
'Dim R22 As New ADODB.Recordset
'
'
'            Open "C:\MERCVB\points.upd" For Output As #2
'
'            Open "C:\MERCVB\customer.upd" For Output As #1
'             ' ena eidos
'              R22.Open "SELECT  * FROM EID  where  KOD LIKE '913%'  ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
'
'             ' ola ta eidh
'             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
'                'On Error GoTo iliadis_error
'
'
'Dim mLTI5 As Single
'Dim MONO As String
'Dim mkod As String
'Dim mbar As String
'Dim MFPA As Integer
'Dim MBARC As String
'
'Dim cp As String
'Dim nc As Long
'
'
'
'
'            Do While Not R22.EOF
'                CC = Space(250)
'                cp = Space(89)
'
'
'                ' Open "C:\POSFILE.TXT" For Output As #1
'
'             If IsNull(R22("kod")) Then
'                mkod = "."
'             Else
'                mkod = R22("kod")
'             End If
'
'
''              If mkod = "005.429" Then
''                 nc = nc + 1
''              End If
'
'             mkod = Replace(mkod, Chr(10), " ")
'             mkod = Replace(mkod, Chr(13), " ")
'
'
'              Mid(CC, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              '  Mid(CC, 3, 15) = Left(mkod + Space(16), 15)   '      Left(mID(mkod, 8, 6) + Space(15), 15)
'              Mid(CC, 3, 15) = Left(mkod + Space(15), 15)
'
'              Mid(CC, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'
'
'              Mid(CC, 246, 2) = "00"
'              Mid(CC, 244, 1) = "1"
'              Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'              Mid(cp, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              Mid(cp, 3, 15) = Left(mkod + Space(15), 15)
'              Mid(cp, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'              Mid(cp, 60, 1) = "0"
'
'              If IsNull(R22("pontoi")) Then
'
'                  Mid(cp, 66, 5) = "0"
'              Else
'                  Mid(cp, 66, 5) = Right("      " + Str(R22("pontoi")), 5)
'
'              End If
'
'              'Mid(cp, 66, 5) = "0"
'
'
'              'Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'
'
' '             0=0   3->15  kodikos  customer.upd   data
''60=>0      66 =>5
'
'
'
'
'
'
'             DoEvents
'
'             Me.Caption = nc
'             nc = nc + 1
'
'
'
'            ' If Len(CC) < 140 Then
'                 Print #2, cp
'             'Else
'                 Print #1, CC
'             'End If
'
'
'                R22.MoveNext
'
'
'                ' If nc > 10 Then Exit Do
'
'            Loop
'
'
'
'
'
'
'             Close #1
'
'             Close #2
'
'
'
'             R22.Close
'
'  Dim ANS As Integer
'
''ANS = MsgBox("Ì· ·ÔÛÙ·ÎÔ˝Ì ÛÙ· pos? ", vbYesNo)
''If ANS = vbYes Then
'
'
' ' If Len(Trim(F_POS1_FOLDER)) > 1 Then
'  '                   FileCopy "C:\MERCVB\Sub TIMES_UPD()
'        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then
'
''OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
''OK----01       036         30                                    –≈—…√—¡÷«                                                              G
''OK----03       087         08                                    ‘…Ã«                                                                D,###.00
''OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
''OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
''OK----17       095         01                                    ‘Ã«Ã¡                                                              I
''----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
''----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
''----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
''OK  ----25       00115                                    ¬  Ÿƒ… œ”
'
'
'
'
'Dim R22 As New ADODB.Recordset
'
'
'            Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2
'
'            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
'             ' ena eidos
'              R22.Open "SELECT  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%' ) ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
'
'             ' ola ta eidh
'             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
'                'On Error GoTo iliadis_error
'
'
'Dim mLTI5 As Single
'Dim MONO As String
'Dim mkod As String
'Dim mbar As String
'Dim MFPA As Integer
'Dim MBARC As String
'
'
'Dim nc As Long
'
'
'            Do While Not R22.EOF
'                ' Open "C:\POSFILE.TXT" For Output As #1
'
'             If IsNull(R22(0)) Then
'                mkod = "."
'             Else
'                mkod = R22(0)
'             End If
'
'
'              If mkod = "005.429" Then
'                 nc = nc + 1
'              End If
'
'             mkod = Replace(mkod, Chr(10), " ")
'             mkod = Replace(mkod, Chr(13), " ")
'
'
'             CC = Left(mkod + Space(21), 20)  ' KVDIKOS
'
'
'
'             'R22 ("ONO")
'             If IsNull(R22("ono")) Then
'                MONO = "."
'             Else
'                MONO = R22("ono")
'             End If
'             MONO = Replace(MONO, Chr(10), " ")
'             MONO = Replace(MONO, Chr(13), " ")
'
'             'R22 ("ONO")
'             If IsNull(R22(1)) Then
'                MBARC = "0000"
'             Else
'                MBARC = R22(1)
'             End If
'
'             MBARC = Replace(MBARC, Chr(10), " ")
'             MBARC = Replace(MBARC, Chr(13), " ")
'
'
'
'             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
'
'
'
'
'
'             CC = CC + Mid(MONO + String(30, " "), 1, 30) + String(21, " ") ' Space(21)
'             If IsNull(R22("LTI5")) Then
'                mLTI5 = 0
'             Else
'                mLTI5 = R22("LTI5")
'             End If
'
'             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
'
'                '  If IsNull(R22("mon")) Then
'                '     CC = CC + Space(3)
'                '   Else
'                '      CC = CC + R22("mon") '+ Space(3), 3)
'                ' End If
'
'
'
'             If IsNull(R22("FPA1")) Then
'                MFPA = 3
'             Else
'                MFPA = R22("FPA1")
'             End If
'
'
'             DoEvents
'
'             Me.Caption = nc
'             nc = nc + 1
'
'             CC = CC + Format(MFPA, "0") + Space(10)
'             CC = CC + " 1"    ' monada Format(R22("mon")
'
'
'   '          On Error Resume Next
'
'            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
'
'             CC = CC + "                       00    0  0      "   ' ÔÌÙÔÈ Êı„ÈÊÔÏÂÌ·
'             CC = Replace(CC, Chr(13), "")
'             If Len(CC) < 140 Then
'                 Print #2, CC
'             Else
'                 Print #1, CC
'             End If
''123456789012340078900230567890
'
'                R22.MoveNext
'            Loop
'
'
'
'
'
'
'             Close #1
'
'             Close #2
'
'
'
'             R22.Close
'
'
'             On Error Resume Next
'             'Dim ANS As Integer
'             'ANS = MsgBox("Ì· ·ÔÛÙ·ÎÔ˝Ì ÛÙ· pos? ", vbYesNo)
'             'If ANS = vbYes Then
'              '   If Len(Trim(F_POS1_FOLDER)) > 1 Then
'               '      FileCopy "C:\MERCVB\POSFILE.TXT", F_POS1_FOLDER + "\POSFILE.TXT"
'                ' Else
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
'                     FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS3\TEC_POS\DATA\POSFILE.TXT"
'                ' End If
'            ' End If
'
'
'' LOAD_PELATES
'
'
'
'
'                On Error Resume Next
'
'                Exit Sub
'
'         '   End If
'
'End Sub
' Sub OLD2LOAD_PELATES()
'        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then
'
''OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
''OK----01       036         30                                    –≈—…√—¡÷«                                                              G
''OK----03       087         08                                    ‘…Ã«                                                                D,###.00
''OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
''OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
''OK----17       095         01                                    ‘Ã«Ã¡                                                              I
''----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
''----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
''----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
''OK  ----25       00115                                    ¬  Ÿƒ… œ”
'
'
'
'
'Dim R22 As New ADODB.Recordset
'
'
'            Open "C:\MERCVB\points.upd" For Output As #2
'
'            Open "C:\MERCVB\customer.upd" For Output As #1
'             ' ena eidos
'              R22.Open "SELECT  * FROM EID  where  KOD LIKE '913%'  ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
'
'             ' ola ta eidh
'             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
'                'On Error GoTo iliadis_error
'
'
'Dim mLTI5 As Single
'Dim MONO As String
'Dim mkod As String
'Dim mbar As String
'Dim MFPA As Integer
'Dim MBARC As String
'
'Dim cp As String
'Dim nc As Long
'
'
'
'
'            Do While Not R22.EOF
'                CC = Space(250)
'                cp = Space(89)
'
'
'                ' Open "C:\POSFILE.TXT" For Output As #1
'
'             If IsNull(R22("kod")) Then
'                mkod = "."
'             Else
'                mkod = R22("kod")
'             End If
'
'
''              If mkod = "005.429" Then
''                 nc = nc + 1
''              End If
'
'             mkod = Replace(mkod, Chr(10), " ")
'             mkod = Replace(mkod, Chr(13), " ")
'
'
'              Mid(CC, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              '  Mid(CC, 3, 15) = Left(mkod + Space(16), 15)   '      Left(mID(mkod, 8, 6) + Space(15), 15)
'              Mid(CC, 3, 15) = Left(mkod + Space(15), 15)
'
'              Mid(CC, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'
'
'              Mid(CC, 246, 2) = "00"
'              Mid(CC, 244, 1) = "1"
'              Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'              Mid(cp, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              Mid(cp, 3, 15) = Left(mkod + Space(15), 15)
'              Mid(cp, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'              Mid(cp, 60, 1) = "0"
'
'              If IsNull(R22("pontoi")) Then
'
'                  Mid(cp, 66, 5) = "0"
'              Else
'                  Mid(cp, 66, 5) = Right("      " + Str(R22("pontoi")), 5)
'
'              End If
'
'              'Mid(cp, 66, 5) = "0"
'
'
'              'Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'
'
' '             0=0   3->15  kodikos  customer.upd   data
''60=>0      66 =>5
'
'
'
'
'
'
'             DoEvents
'
'             Me.Caption = nc
'             nc = nc + 1
'
'
'
'            ' If Len(CC) < 140 Then
'                 Print #2, cp
'             'Else
'                 Print #1, CC
'             'End If
'
'
'                R22.MoveNext
'
'
'                ' If nc > 10 Then Exit Do
'
'            Loop
'
'
'
'
'
'
'             Close #1
'
'             Close #2
'
'
'
'             R22.Close
'
'
'
'                     FileCopy "C:\MERCVB\customer.upd", "\\POS1\TEC_POS\DATA\customer.upd"
'                     FileCopy "C:\MERCVB\customer.upd", "\\POS2\TEC_POS\DATA\customer.upd"
'                     FileCopy "C:\MERCVB\customer.upd", "\\POS3\TEC_POS\DATA\customer.upd"
'
'                     FileCopy "C:\MERCVB\points.upd", "\\POS1\TEC_POS\DATA\points.upd"
'                      FileCopy "C:\MERCVB\points.upd", "\\POS2\TEC_POS\data\points.upd"
'                     FileCopy "C:\MERCVB\points.upd", "\\POS3\TEC_POS\data\points.upd"
'
'
'End Sub






Sub OLDTIMES_UPD()
        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    –≈—…√—¡÷«                                                              G
'OK----03       087         08                                    ‘…Ã«                                                                D,###.00
'OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
'OK----17       095         01                                    ‘Ã«Ã¡                                                              I
'----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
'----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
'----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
'OK  ----25       00115                                    ¬  Ÿƒ… œ”


Dim RPOS1 As New ADODB.Recordset
Dim WRITEOK As Integer
Dim R22 As New ADODB.Recordset
Dim SQL1 As String, K1 As Integer
Dim CC As String



Dim ARR1(10) As String

           ' Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2

            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
             ' ena eidos
              R22.Open "SELECT  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%' ) ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                'On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim mkod As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String


Dim nc As Long


            Do While Not R22.EOF
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22(0)) Then
                mkod = "."
             Else
                mkod = R22(0)
             End If
                
                
              If mkod = "005.429" Then
                 nc = nc + 1
              End If
                
             mkod = Replace(mkod, Chr(10), " ")
             mkod = Replace(mkod, Chr(13), " ")
                
                
             CC = Left(mkod + Space(21), 20)  ' KVDIKOS
             
             
             
             'R22 ("ONO")
             If IsNull(R22("ono")) Then
                MONO = "."
             Else
                MONO = R22("ono")
             End If
             MONO = Replace(MONO, Chr(10), " ")
             MONO = Replace(MONO, Chr(13), " ")
             
             'R22 ("ONO")
             If IsNull(R22(1)) Then
                MBARC = "0000"
             Else
                MBARC = R22(1)
             End If
             
             MBARC = Replace(MBARC, Chr(10), " ")
             MBARC = Replace(MBARC, Chr(13), " ")

             
             
             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
             
             
             
             
             
             CC = CC + Mid(MONO + String(30, " "), 1, 30) + String(21, " ") ' Space(21)
             If IsNull(R22("LTI5")) Then
                mLTI5 = 0
             Else
                mLTI5 = R22("LTI5")
             End If
             
             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
     
                '  If IsNull(R22("mon")) Then
                '     CC = CC + Space(3)
                '   Else
                '      CC = CC + R22("mon") '+ Space(3), 3)
                ' End If
     
             
             
             If IsNull(R22("FPA1")) Then
                MFPA = 3
             Else
                MFPA = R22("FPA1")
             End If
             
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             CC = CC + Format(MFPA, "0") + Space(10)
             CC = CC + " 1"    ' monada Format(R22("mon")
             
             
   '          On Error Resume Next
             
            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
             
             CC = CC + "                       00    0  0      "   ' ÔÌÙÔÈ Êı„ÈÊÔÏÂÌ·
             CC = Replace(CC, Chr(13), "")
             'If Len(cc) < 140 Then
               '  Print #2, cc
            ' Else
            
            
            
            
            
            WRITEOK = 0
            '≈À≈√◊Ÿ ¡Õ ’–¡—◊≈… «ƒ« « ≈√√—¡÷« ”‘œ EIDPOS1 KAI AN EINAI IDIA WRITEOK=1 NA GRAFEI  =0 NA MHN GRAFEI
            RPOS1.Open "SELECT * FROM EIDPOS1 WHERE ERG='" + R22!ERG + "'", GDB, adOpenDynamic, adLockOptimistic
            If RPOS1.EOF Then
                WRITEOK = 1
            Else
                For K1 = 0 To 5
                   If RPOS1(K1) = R22(K1) Then
                   Else
                       WRITEOK = 1
                   End If
                Next
            End If
            
            
            If WRITEOK = 1 Then
                If IsNull(R22(0)) Then
                   ARR1(0) = " "
                Else
                   ARR1(0) = R22(0)
                End If
                
                If IsNull(R22(1)) Then
                   ARR1(1) = " "
                Else
                   ARR1(1) = R22(1)
                End If
                
                If IsNull(R22(2)) Then 'ONO
                   ARR1(2) = " "
                Else
                   ARR1(2) = Replace(R22(2), "'", "`")
                End If
                
                If IsNull(R22(3)) Then ' LTI5
                   ARR1(3) = "0"
                Else
                   ARR1(3) = Replace(Str(R22(3)), ",", ".")
                End If
                
                If IsNull(R22(4)) Then 'MON
                   ARR1(4) = " "
                Else
                   ARR1(4) = R22(4)
                End If
                
                 If IsNull(R22(5)) Then ' FPA
                   ARR1(5) = "3"
                Else
                   ARR1(5) = Replace(Str(R22(5)), ",", ".")
                End If
                
                 If IsNull(R22(6)) Then ' TMHMA
                   ARR1(6) = "3"
                Else
                   ARR1(6) = Replace(Str(R22(6)), ",", ".")
                End If
                SQL1 = "INSERT INTO EIDPOS1 ([KOD],[ERG],[ONO],[LTI5],[MON],[FPA1],[TMHMA]) VALUES("
                SQL1 = SQL1 + "'" + ARR1(0) + "',"
                SQL1 = SQL1 + "'" + ARR1(1) + "',"
                SQL1 = SQL1 + "'" + ARR1(2) + "',"
                SQL1 = SQL1 + "" + ARR1(3) + ","
                SQL1 = SQL1 + "'" + ARR1(4) + "',"
                SQL1 = SQL1 + "" + ARR1(5) + ","  'FPA
                SQL1 = SQL1 + "" + ARR1(6) + " )"  'TMHMA
                GDB.Execute SQL1
                Print #1, CC
                
            End If
            
            
            RPOS1.Close
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
            
               
            
            
            
            
            
            ' End If
'123456789012340078900230567890
                  
                R22.MoveNext
            Loop






             Close #1
             
            ' Close #2
             
             
             
             R22.Close
             
             
             On Error Resume Next
             

FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS1\TEC_POS\DATA\POSFILE.TXT"
FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
FileCopy "C:\MERCVB\POSFILE.TXT", "\\POS3\TEC_POS\DATA\POSFILE.TXT"

LOAD_PELATES




                On Error Resume Next
                
                Exit Sub

         '   End If

End Sub
'Sub OLDLOAD_PELATES()
'        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then
'
''OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
''OK----01       036         30                                    –≈—…√—¡÷«                                                              G
''OK----03       087         08                                    ‘…Ã«                                                                D,###.00
''OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
''OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
''OK----17       095         01                                    ‘Ã«Ã¡                                                              I
''----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
''----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
''----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
''OK  ----25       00115                                    ¬  Ÿƒ… œ”
'
'
'
'
'Dim R22 As New ADODB.Recordset
'
'
'            Open "C:\MERCVB\points.upd" For Output As #2
'
'            Open "C:\MERCVB\customer.upd" For Output As #1
'             ' ena eidos
'              R22.Open "SELECT  * FROM EID  where  KOD LIKE '913%'  ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
'
'             ' ola ta eidh
'             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
'                'On Error GoTo iliadis_error
'
'
'Dim mLTI5 As Single
'Dim MONO As String
'Dim mkod As String
'Dim mbar As String
'Dim MFPA As Integer
'Dim MBARC As String
'
'Dim cp As String
'Dim nc As Long
'Dim CC As String
'
'
'
'
'            Do While Not R22.EOF
'                CC = Space(250)
'                cp = Space(89)
'
'
'                ' Open "C:\POSFILE.TXT" For Output As #1
'
'             If IsNull(R22("kod")) Then
'                mkod = "."
'             Else
'                mkod = R22("kod")
'             End If
'
'
''              If mkod = "005.429" Then
''                 nc = nc + 1
''              End If
'
'             mkod = Replace(mkod, Chr(10), " ")
'             mkod = Replace(mkod, Chr(13), " ")
'
'
'              Mid(CC, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              '  Mid(CC, 3, 15) = Left(mkod + Space(16), 15)   '      Left(mID(mkod, 8, 6) + Space(15), 15)
'              Mid(CC, 3, 15) = Left(mkod + Space(15), 15)
'
'              Mid(CC, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'
'
'              Mid(CC, 246, 2) = "00"
'              Mid(CC, 244, 1) = "1"
'              Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'              Mid(cp, 1, 1) = "0"  'Left(mkod + Space(15), 15)
'              Mid(cp, 3, 15) = Left(mkod + Space(15), 15)
'              Mid(cp, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
'
'              Mid(cp, 60, 1) = "0"
'
'              If IsNull(R22("pontoi")) Then
'
'                  Mid(cp, 66, 5) = "0"
'              Else
'                  Mid(cp, 66, 5) = Right("      " + Str(R22("pontoi")), 5)
'
'              End If
'
'              'Mid(cp, 66, 5) = "0"
'
'
'              'Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
'
'
'
'
'
'
' '             0=0   3->15  kodikos  customer.upd   data
''60=>0      66 =>5
'
'
'
'
'
'
'             DoEvents
'
'             Me.Caption = nc
'             nc = nc + 1
'
'
'
'            ' If Len(CC) < 140 Then
'                 Print #2, cp
'             'Else
'                 Print #1, CC
'             'End If
'
'
'                R22.MoveNext
'
'
'                ' If nc > 10 Then Exit Do
'
'            Loop
'
'
'
'
'
'
'             Close #1
'
'             Close #2
'
'
'
'             R22.Close
'
'
'             On Error Resume Next
'             Exit Sub
'
'Exit Sub
'
'
'FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS1\TEC_POS\FILES\UPDATE.CUS"
'FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS2\TEC_POS\FILES\UPDATE.CUS"
'FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS3\TEC_POS\FILES\UPDATE.CUS"
'
'
'
'FileCopy "C:\MERCVB\points.upd", "\\POS1\TEC_POS\FILES\UPDATE.CUS"
'FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS2\TEC_POS\FILES\UPDATE.CUS"
'FileCopy "C:\MERCVB\UPDATE.CUS", "\\POS3\TEC_POS\FILES\UPDATE.CUS"
'                On Error Resume Next
'
'                Exit Sub
'
'         '   End If
'
'End Sub

Private Sub EID_Click()
  
        Dim K

        Dim fname

        Dim z

510     On Error GoTo SEELINE

520     F_RUNNING = 1

        Dim RR       As New ADODB.Recordset

        Dim RL       As New ADODB.Recordset

        Dim r2       As New ADODB.Recordset

        Dim rALLAGES As New ADODB.Recordset

        Dim CC

        Dim sql As String

        Dim J   As Long

        Dim mK  As String

        'GDB.Execute "DELETE FROM EID"
        ' RR.Open "SELECT TOP 10 * FROM EID", GDBR, adOpenDynamic, adLockOptimistic

        '530   RL.Open "SELECT TOP 10 * FROM EID", GDB, adOpenDynamic, adLockOptimistic

        'ƒ…¡¬¡∆Ÿ ‘…” ¡ÀÀ¡√≈” –œ’ ≈√…Õ¡Õ  ACTION = INS , DEL , UPD  OTAN ENHMERVNONTAI GINONTA 1INS,1UPD,1DEL

              List1.BackColor = lbl≈…ƒ«¡–œ.BackColor


        rALLAGES.Open "SELECT count(*) FROM EIDALLAGES WHERE (NOT KOD IS NULL) AND LEFT(ACTION,1) IN ('I','D','U')  and " + FTAM_FIELD + " is NULL ", GDBR, adOpenDynamic, adLockOptimistic
        List1.AddItem Str(rALLAGES(0)) + " AÀÀ¡√≈” ≈…ƒŸÕ"

        J = rALLAGES(0)
        rALLAGES.Close

540     rALLAGES.Open "SELECT TOP 50 * FROM EIDALLAGES WHERE (NOT KOD IS NULL) AND LEFT(ACTION,1) IN ('I','D','U') and " + FTAM_FIELD + " is NULL ORDER BY ID ", GDBR, adOpenDynamic, adLockOptimistic

550     'On Error GoTo 0
        Dim C

560     z = 0

570     Do While Not rALLAGES.EOF
          
580         DoEvents
590         mK = rALLAGES("kod")

            '----------------------------------------------------------------------------------------------
600         If Left(rALLAGES("ACTION"), 3) = "DEL" Then
610             List1.AddItem "DEL-----------"
620             GDB.Execute "DELETE FROM EID WHERE KOD='" + mK + "'"
        
                '-----------------------------------------------------------------------------------------------------------------
630         Else

640             If Left(rALLAGES("ACTION"), 3) = "INS" Then
650                 r2.Open "SELECT COUNT(*) FROM EID where KOD='" + Trim(mK) + "';", GDB, adOpenDynamic, adLockOptimistic

660
                    '·Ì ‰ÂÌ ı·Ò˜ÂÈ ÙÔ ·ÌÔÈ„˘ ÙÔÈÍ·
                    If r2(0) = 0 Then
670                     GDB.Execute "INSERT INTO EID (KOD) VALUES ('" + mK + "')"
680                 End If

690                 r2.Close
700             End If
          
                'RR.Close
710             List1.AddItem "≈…ƒœ”=" + mK

720             RR.Open "SELECT * FROM EID WHERE KOD='" + Trim(mK) + "'", GDBR, adOpenDynamic, adLockOptimistic
                f_ExistOnServer = 1

                If RR.EOF Then
                    GDBR.Execute "INSERT INTO EID (KOD) VALUES ('" + Trim(mK) + "')"
                    'DEN YPARXEI STHN MANA
       
                    f_ExistOnServer = 0
                    GoTo 1780
                End If

                On Error Resume Next

                fRL.Close

                On Error GoTo SEELINE

                '·ÌÔÈ„˘ Í·È ÙÔ ÂÈ‰ÔÚ ÙÔÈÍ· „È· Ì· ‚ÎÂ˘ ÙÈÚ ‰È·ˆÔÒ›Ú
                fRL.Open "SELECT * FROM EID WHERE KOD='" + mK + "'", GDB, adOpenDynamic, adLockOptimistic

                F_SQL = ""

                If Not IsNull(RR("ONO")) Then
                    UPDSTR "ONO", Left(RR("ONO"), 35), mK
                End If
          
                '230       C = RR("ONO")
                '240       If Not IsNull(C) Then
                '250          GDB.Execute "UPDATE EID SET ONO='" + C + "' where KOD='" + mK + _
                '    "';"
                '260       End If
          
740             UPDSTR "ONO2", RR("ONO2"), mK
750             UPDSTR "ERG", RR("ERG"), mK
760             UPDSTR "KODSYNOD", RR("KODSYNOD"), mK

780             UPDSTR "MON", RR("MON"), mK
790             UPDSTR "KODLOG", RR("KODLOG"), mK
800             UPDSTR "KODLOGAG", RR("KODLOGAG"), mK
          
810             UPDSTR "PROM", RR("PROM"), mK
820             UPDSTR "MEMO", RR("MEMO"), mK
          
830             UPDSTR "NUM1", RR("NUM1"), mK
840             UPDSTR "NUM2", RR("NUM2"), mK
850             UPDSTR "NUM3", RR("NUM3"), mK
          
860             UPDSTR "CH1", RR("CH1"), mK
870             UPDSTR "CH2", RR("CH2"), mK
880             UPDSTR "CH3", RR("CH3"), mK
890             UPDSTR "CH4", RR("CH4"), mK
900             UPDSTR "CH5", RR("CH5"), mK
910             UPDSTR "CH6", RR("CH6"), mK
          
920             UPDSTR "KODERG", RR("KODERG"), mK
930             UPDSTR "UES", RR("UES"), mK
          
940             UPDNUM "LITRA", RR("LITRA"), mK
950             UPDNUM "XTI", RR("XTI"), mK
960             UPDNUM "LTI", RR("LTI"), mK
970             UPDNUM "LTI2", RR("LTI2"), mK
980             UPDNUM "LTI3", RR("LTI3"), mK
990             UPDNUM "LTI4", RR("LTI4"), mK
1000            UPDNUM "LTI5", RR("LTI5"), mK
1010            UPDNUM "APOS", RR("APOS"), mK
1020            UPDNUM "EISSYN", RR("EISSYN"), mK
1030            UPDNUM "EXSYN", RR("EXSYN"), mK
1040            UPDNUM "EXMHN", RR("EXMHN"), mK
1050            UPDNUM "POS", RR("POS"), mK
1060            UPDNUM "EISMHN", RR("EISMHN"), mK
1070            UPDNUM "DESMIA", RR("DESMIA"), mK

1080            UPDNUM "APOS01", RR("APOS01"), mK
1090            UPDNUM "APOS02", RR("APOS02"), mK
1100            UPDNUM "APOS03", RR("APOS03"), mK
1110            UPDNUM "APOS04", RR("APOS04"), mK

1120            UPDNUM "POS01", RR("POS01"), mK
1130            UPDNUM "POS02", RR("POS02"), mK
1140            UPDNUM "POS03", RR("POS03"), mK
1150            UPDNUM "POS04", RR("POS04"), mK

1160            UPDNUM "SPA", RR("SPA"), mK
1170            UPDNUM "FPA", RR("FPA"), mK
1180            UPDNUM "PAR", RR("PAR"), mK
          
1190            UPDNUM "AEG", RR("AEG"), mK
1200            UPDNUM "M01", RR("M01"), mK
1210            UPDNUM "M02", RR("M02"), mK
1220            UPDNUM "M03", RR("M03"), mK
1230            UPDNUM "M04", RR("M04"), mK
1240            UPDNUM "M05", RR("M05"), mK
1250            UPDNUM "M06", RR("M06"), mK
1260            UPDNUM "M07", RR("M07"), mK
1270            UPDNUM "M08", RR("M08"), mK
1280            UPDNUM "M09", RR("M09"), mK
1290            UPDNUM "M10", RR("M10"), mK
1300            UPDNUM "M11", RR("M11"), mK
1310            UPDNUM "M12", RR("M12"), mK

1320            UPDNUM "G01", RR("G01"), mK
1330            UPDNUM "G02", RR("G02"), mK
1340            UPDNUM "G03", RR("G03"), mK
1350            UPDNUM "G04", RR("G04"), mK
1360            UPDNUM "G05", RR("G05"), mK
1370            UPDNUM "G06", RR("G06"), mK
1380            UPDNUM "G07", RR("G07"), mK
1390            UPDNUM "G08", RR("G08"), mK
1400            UPDNUM "G09", RR("G09"), mK
1410            UPDNUM "G10", RR("G10"), mK
1420            UPDNUM "G11", RR("G11"), mK
1430            UPDNUM "G12", RR("G12"), mK
1440            UPDNUM "POS_KERD2", RR("POS_KERD2"), mK
1450            UPDNUM "POS_KERD", RR("POS_KERD"), mK
1460            UPDNUM "POS_KERD3", RR("POS_KERD3"), mK

1470            UPDNUM "POS_EKPT", RR("POS_EKPT"), mK
1480            UPDNUM "KODFIAL", RR("KODFIAL"), mK
1490            UPDNUM "POS_EKPT", RR("POS_EKPT"), mK
1500            UPDNUM "AJIAPO", RR("AJIAPO"), mK
1510            UPDNUM "AJIPOL", RR("AJIPOL"), mK
1520            UPDNUM "AJIAGO", RR("AJIAGO"), mK
1530            UPDNUM "MESXTI", RR("MESXTI"), mK
1540            UPDNUM "LIT_SYNOD", RR("LIT_SYNOD"), mK
1550            UPDNUM "LTI_SYNOD", RR("LTI_SYNOD"), mK

1560            UPDNUM "SYSKEYASIA", RR("SYSKEYASIA"), mK
1570            UPDNUM "PROMHU", RR("PROMHU"), mK

1580            UPDNUM "SYNAGO", RR("SYNAGO"), mK
1590            UPDNUM "SYNPOL", RR("SYNPOL"), mK

1600            UPDNUM "EISSYN0", RR("EISSYN0"), mK
1610            UPDNUM "EXSYN0", RR("EXSYN0"), mK

1620            UPDNUM "SYNAGO0", RR("SYNAGO0"), mK
1630            UPDNUM "SYNPOL0", RR("SYNPOL0"), mK

1640            UPDNUM "EISMHN0", RR("EISMHN0"), mK
1650            UPDNUM "EXMHN0", RR("EXMHN0"), mK

1700            UPDNUM "PONTOI", RR("PONTOI"), mK
1710            UPDNUM "EPIUYP", RR("EPIUYP"), mK

1720            UPDNUM "SYSKMAX", RR("SYSKMAX"), mK
1730            UPDNUM "SYSKMIN", RR("SYSKMIN"), mK

1740            UPDNUM "MESTIMPOL", RR("MESTIMPOL"), mK

1750            UPDDAT "HM1", RR("HM1"), mK
1760            UPDDAT "HM2", RR("HM2"), mK
1770            UPDDAT "HM3", RR("HM3"), mK
                ' UPDDAT "LASTUPD", RR("LASTUPD"), mK
          
                '     GDB.Execute "UPDATE EID SET " + Left(F_SQL, Len(F_SQL) - 1) + " where KOD='" + mK + "'", K
          
1780            RR.Close
       
                '=========================================================================================================
             
1790        End If

            '1800    GDBR.Execute _
            '         "UPDATE EIDALLAGES SET ACTION='1'+LEFT(ACTION,3) WHERE KOD='" + rALLAGES("kod") _
            '        + "'"
          
1800        GDBR.Execute "UPDATE EIDALLAGES SET " + FTAM_FIELD + "=1 WHERE KOD='" + rALLAGES("kod") + "'"
          
1810        z = z + 1
            ' If z Mod 10 = 0 Then
1820        Me.Caption = Format(z, "####") + "/" + Str(J) + " " + rALLAGES("KOD")
            ' End If
1830        List1.AddItem Format(z, "####") + "/" + Str(J) + "  " + rALLAGES("KOD")
        
1840        rALLAGES.MoveNext

1850        DoEvents

1860    Loop

1870    If List1.ListCount > 30 Then List1.Clear

1880    List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + " ≈Õ«Ã≈—Ÿ»« ¡Õ " + Str(z) + " ≈…ƒ« "

1890    F_RUNNING = 0

1900    Exit Sub

SEELINE:

        List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"

        '      Resume Next
        On Error Resume Next

        GDB.Close
        GDBR.Close
        Me.Caption = Str(Erl) + "-----" + Err.Description
        SAVE_ERROR Err.Description & " in EID_Click " & "at line " & Erl
        open_data

        Exit Sub    'Resume Next

End Sub
Private Sub EnhmBarcodes()
      Dim K
      Dim fname
      Dim z


2010  On Error GoTo SEELINE

2020  F_RUNNING = 1

      Dim RR As New ADODB.Recordset
      Dim RL As New ADODB.Recordset
      Dim r2 As New ADODB.Recordset
      Dim rALLAGES As New ADODB.Recordset


      Dim CC
      Dim sql As String
      Dim J As Long
      
      List1.BackColor = lbl≈…ƒ«¡–œ.BackColor

      'GDB.Execute "DELETE FROM EID"
      ' RR.Open "SELECT TOP 10 * FROM EID", GDBR, adOpenDynamic, adLockOptimistic

2030  RL.Open "SELECT TOP 10 * FROM BARCODES", GDB, adOpenDynamic, _
          adLockOptimistic

      'ƒ…¡¬¡∆Ÿ ‘…” ¡ÀÀ¡√≈” –œ’ ≈√…Õ¡Õ  ACTION = INS , DEL , UPD  OTAN ENHMERVNONTAI GINONTA 1INS,1UPD,1DEL
2040  rALLAGES.Open _
          "SELECT * FROM BARCODESALLAGES WHERE NOT ERG IS NULL AND  LEFT(ACTION,1) IN ('I','D','U') and " + FTAM_FIELD + " is NULL ORDER BY ID ", _
          GDBR, adOpenDynamic, adLockOptimistic

2050  z = 0
2060  Do While Not rALLAGES.EOF
          
2070    DoEvents
        ' On Error GoTo 0
2080    If Left(rALLAGES("ACTION"), 3) = "DEL" And Not IsNull(rALLAGES("ERG")) Then
2090       List1.AddItem "DEL BARCODE--------"
2100       sql = "DELETE FROM BARCODES WHERE ERG='" + Trim(rALLAGES("ERG")) + "'"
2110       List1.AddItem sql
2120       GDB.Execute sql
2130    List1.AddItem "OK " + sql
        
2140    ElseIf Left(rALLAGES("ACTION"), 3) = "INS" Then
2150       List1.AddItem "INS-----------"
           ';RR.Open "SELECT * FROM BARCODES where KOD='" + rALLAGES("KOD") + "';", GDBR, adOpenDynamic, adLockOptimistic

            'Ã«Õ ‘’◊œÕ  ¡… ’–¡—◊≈… «ƒ«
2170         r2.Open "SELECT COUNT(*) FROM BARCODES where ERG='" + _
          rALLAGES("ERG") + "';", GDB, adOpenDynamic, adLockOptimistic
2180         If r2(0) = 0 Then
2190             GDB.Execute "INSERT INTO BARCODES (ERG,KOD) VALUES ('" + _
          rALLAGES("ERG") + "','" + rALLAGES("KOD") + "')"
2200         End If
2210         r2.Close
2220        ' RR.Close
             
2230    ElseIf Left(rALLAGES("ACTION"), 3) = "UPD" Then
2240        List1.AddItem "UPD BARCODES???"
2250    End If
      '  List1.AddItem "UPDate barcodesallages"
    GDBR.Execute _
          "UPDATE BARCODESALLAGES SET " + FTAM_FIELD + "=1 WHERE ERG='" + _
          rALLAGES("ERG") + "' and ID=" + Str(rALLAGES("ID"))
2270        z = z + 1
           ' If z Mod 10 = 0 Then
2280           Me.Caption = rALLAGES("ERG")
           ' End If


        
2290        rALLAGES.MoveNext
2300        DoEvents



2310      Loop


2320      If List1.ListCount > 30 Then List1.Clear

2330     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + " ≈Õ«Ã≈—Ÿ»« ¡Õ " + _
          Str(z) + " BARCODES "
         

2340  F_RUNNING = 0

2350  Exit Sub

SEELINE:
          'HandleError "MdiForm-load"
          'Resume Next
2360    '      MsgBox fname + " - " + Str(RR(K).Type) + Str(Erl) + " - " + _
         ' Err.Description
2370      List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"
2380      On Error Resume Next
2390      GDB.Close
2400      GDBR.Close
SAVE_ERROR Err.Description & " in ENHMBARCODES_Click " & "at line " & Erl
2410      Me.Caption = Str(Erl) + "-----" + Err.Description
2420 ' MsgBox Str(Erl) + Err.Description
2430  'End


2440      open_data
          


2450      Exit Sub    'Resume Next

End Sub

Private Sub DELEGGTIM()
      Dim K
      Dim fname
      Dim z


2460  On Error GoTo SEELINE

2470  F_RUNNING = 1

      Dim RR As New ADODB.Recordset
      Dim RL As New ADODB.Recordset
      Dim r2 As New ADODB.Recordset
      Dim rALLAGES As New ADODB.Recordset


      Dim CC
      Dim sql As String
      Dim J As Long


    Dim N As Long
    Dim ID As Long
    

RL.Open "select * from EGGTIMDEL", GDB, adOpenDynamic, adLockOptimistic

'If RL(0) >= 1 Then
 '   GDB.Execute "DELETE FROM  [PLATEIA].[MERCURY].dbo.EGGTIM WHERE ID>0 AND APOT=3  AND ID IN  (SELECT ID FROM EGGTIMDEL)", N
  Do While Not RL.EOF
      If Not IsNull(RL("ID")) Then
         ID = RL("ID")
         If ID > 0 Then
            GDBR.Execute "DELETE FROM EGGTIM WHERE ID=" + Str(ID), N
            If N > 0 Then
               GDB.Execute "DELETE FROM EGGTIMDEL WHERE ID=" + Str(ID)
               List1.AddItem "**** ƒ…≈√—¡÷« ‘œ ID " + Str(ID)
            End If
         End If
      End If
      RL.MoveNext
  Loop
  
   
Exit Sub











SEELINE:
          'HandleError "MdiForm-load"
          'Resume Next
2680     Resume Next
'MsgBox fname + " - " + Str(RR(K).Type) + Str(Erl) + " - " + _
          Err.Description
2690      List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"
2700      On Error Resume Next
2710      GDB.Close
2720      GDBR.Close
          SAVE_ERROR Err.Description & " in DELEGGTIM_Click " & "at line " & Erl

2730      Me.Caption = Str(Erl) + "-----" + Err.Description
2740  MsgBox Str(Erl) + Err.Description
2750  End


2760      open_data



2770      Exit Sub    'Resume Next

End Sub










'Private Sub DELEGGTIM()
'      Dim K
'      Dim fname
'      Dim z
'
'
'2460  On Error GoTo SEELINE
'
'2470  F_RUNNING = 1
'
'      Dim RR As New ADODB.Recordset
'      Dim RL As New ADODB.Recordset
'      Dim r2 As New ADODB.Recordset
'      Dim rALLAGES As New ADODB.Recordset
'
'
'      Dim CC
'      Dim SQL As String
'      Dim J As Long
'
'      'GDB.Execute "DELETE FROM EID"
'      ' RR.Open "SELECT TOP 10 * FROM EID", GDBR, adOpenDynamic, adLockOptimistic
'Dim MATIM As String, MDAT As String
'
'
'      'ƒ…¡¬¡∆Ÿ ‘…” ¡ÀÀ¡√≈” –œ’ ≈√…Õ¡Õ  ACTION = INS , DEL , UPD  OTAN ENHMERVNONTAI GINONTA 1INS,1UPD,1DEL
'2480  rALLAGES.Open _
'          "SELECT * FROM EGGTIMALLAGES2 WHERE IDE=0 AND ID>0 ", _
'          GDB, adOpenDynamic, adLockOptimistic
'
'2490  z = 0
'2500  Do While Not rALLAGES.EOF
'
'2510    DoEvents
'            z = z + 1
'2530       List1.AddItem "DEL EGGTIM--------"
'          SQL = "DELETE FROM EGGTIM WHERE KOLA=3 AND ID=" + Str(rALLAGES("ID"))
'               List1.AddItem SQL
'          GDBR.Execute SQL, K
'
'2590     GDB.Execute _
'            "UPDATE EGGTIMALLAGES2 SET ID=0 WHERE ID='" + _
'            Str(rALLAGES("ID")) + "'"
'2610        rALLAGES.MoveNext
'
'2630      Loop
'
'
'2640      If List1.ListCount > 30 Then List1.Clear
'
'2650     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + " ≈Õ«Ã≈—Ÿ»« ¡Õ " + _
'          Str(z) + " BARCODES "
'         rALLAGES.Close
'
'
'2660  F_RUNNING = 0
'
'2670  Exit Sub
'
'SEELINE:
'          'HandleError "MdiForm-load"
'          'Resume Next
'2680     Resume Next
'MsgBox fname + " - " + Str(RR(K).Type) + Str(Erl) + " - " + _
'          Err.Description
'2690      List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"
'2700      On Error Resume Next
'2710      GDB.Close
'2720      GDBR.Close
'2730      Me.Caption = Str(Erl) + "-----" + Err.Description
'2740  MsgBox Str(Erl) + Err.Description
'2750  End
'
'
'2760      open_data
'
'
'
'2770      Exit Sub    'Resume Next
'
'
'End Sub

Private Sub Form_DblClick()

100     PARAMETROI.PARAM.Caption = "SYNCHRO"
110     PARAMETROI.Show 1

End Sub

Private Sub Form_Load()

'F_APOT = "4"



If App.PrevInstance Then End
2780    Me.Show

2790    open_data


'F_SYNCHRO_KIN = 1

F_SYNCHRO_KIN = Val(FINDPARAMETROI(1, "SYNCHRO", "F_SYNCHRO_KIN", "0", "¡–œ”‘œÀ«  …Õ«”≈ŸÕ ”≈  ≈Õ‘—… œ "))
F_POS = Val(FINDPARAMETROI(1, "SYNCHRO", "F_POS", "0", "≈NHME—Ÿ”H ¡–œ POS =1  œ˜È=0 "))
F_PONTOI = Val(FINDPARAMETROI(1, "SYNCHRO", "F_PONTOI", "0", "≈NHME—Ÿ”H –œÕ‘ŸÕ =1  œ˜È=0 "))
F_APOT = (FINDPARAMETROI(1, "SYNCHRO", "F_APOT", "4", "¡—…»Ãœ” ¡–œ»« «” 1-4 "))
F_ENHM_EIDH = Val(FINDPARAMETROI(1, "SYNCHRO", "F_ENHM_EIDH", "0", "≈NHME—Ÿ”H ≈…ƒŸÕ&BARCODES ¡–œ SERVER =1  œ˜È=0 "))
F_MHXANH = FINDPARAMETROI(1, "SYNCHRO", "F_MHXANH", "0", "MHXANH ")
F_LOADPEL = Val(FINDPARAMETROI(1, "SYNCHRO", "F_LOADPEL", "0", "¡–œ”‘œÀ« ‘…ÃŸÕ ”≈ POS =1  œ˜È=0 "))


f_apotL.Caption = "·ÔË " + F_APOT

End Sub

Private Sub open_data()

 

      Dim a
2800  a = 1
2810   F_RUNNING = 1

2820  On Error GoTo topikos

      Dim sServer, sLogin, sPassword
      Dim sServerR, sLoginR, sPasswordR

2830   Open "C:\MERCVB\REMOTE.TXT" For Input As #1

 Line Input #1, FTAM_FIELD  'T30 XANTHI     T40-T43 THESSALONIKI
  Line Input #1, sServer  'TOPIKO DSN
  Line Input #1, sLogin
  Line Input #1, sPassword

  'Line Input #1, sServer2
  'Line Input #1, sLogin2
  'Line Input #1, sPassword2

  'Line Input #1, sServer3
  'Line Input #1, sLogin3
  'Line Input #1, sPassword3

'3 Ã«◊¡Õ«Ã¡‘¡

 Line Input #1, sServerR   'REMOTE DSN
 Line Input #1, sLoginR
 Line Input #1, sPasswordR

2900  Close #1


top:

List1.AddItem "REMOTE.TXT ƒ…¡¬¡”‘« ≈"
lTameio.Caption = FTAM_FIELD

  SERV1 = 0
  Do While True
      If Check.Value = vbChecked Then
         End
      End If
      
      If IsValidODBCLogin(sServer, sLogin, sPassword) = True Then
           ' GDB.Open "DSN=" + sServer + ";UID=" + sLogin + ";PWD=" + _
          'sPassword
          
          gCONNECT = "DSN=" + sServer + ";UID=" + sLogin + ";PWD=" + sPassword
          GDB.Open gCONNECT
          
                List1.AddItem "1 . ‘œ–… « ”’Õƒ≈”« œ "
                SERV1 = 1
            Exit Do
      'MsgBox "Connection Successful", vbInformation, "ODBC Logon"
      Else
                  Me.Caption = "–—œ”–¡»≈…¡ ”’Õƒ≈”«” Ã≈ ‘œ–… œ 1 SERVER " + Format(Now, "HH:MM")
                  DoEvents
          '      MsgBox "Connection Failed", vbExclamation, "ODBC Logon"
      End If
    DoEvents
  Loop

'
'  SERV2 = 0
'  Do While True
'      If IsValidODBCLogin(sServer2, sLogin2, sPassword2) = True Then
'            GDB2.Open "DSN=" + sServer2 + ";UID=" + sLogin2 + ";PWD=" + _
'          sPassword2
'                List1.AddItem "2 . ‘œ–… « ”’Õƒ≈”« œ "
'                SERV2 = 1
'            Exit Do
'      'MsgBox "Connection Successful", vbInformation, "ODBC Logon"
'      Else
'                  Me.Caption = "–—œ”–¡»≈…¡ ”’Õƒ≈”«” Ã≈ ‘œ–… œ 2 SERVER " + Format(Now, "HH:MM")
'                  DoEvents
'          '      MsgBox "Connection Failed", vbExclamation, "ODBC Logon"
'      End If
'
'  Loop
'
'
'
'  SERV3 = 0
'  Do While True
'      If IsValidODBCLogin(sServer3, sLogin3, sPassword3) = True Then
'            GDB3.Open "DSN=" + sServer3 + ";UID=" + sLogin3 + ";PWD=" + _
'          sPassword3
'                List1.AddItem "3 . ‘œ–… « ”’Õƒ≈”« œ "
'                SERV3 = 1
'            Exit Do
'      'MsgBox "Connection Successful", vbInformation, "ODBC Logon"
'      Else
'                  Me.Caption = "–—œ”–¡»≈…¡ ”’Õƒ≈”«” Ã≈ ‘œ–… œ 3 SERVER " + Format(Now, "HH:MM")
'                  DoEvents
'          '      MsgBox "Connection Failed", vbExclamation, "ODBC Logon"
'      End If
'
'  Loop
'
'











2990  Me.Caption = "‘œ–… œ” œ "
3000  a = 1
3010  On Error GoTo remote



      '230   sServerR = "MAGAZI"
      '240   sLoginR = "sa"
      '250   sPasswordR = "epsilonsa"



3020  Do While True
      If Check.Value = vbChecked Then
         End
      End If



3030      If IsValidODBCLogin(sServerR, sLoginR, sPasswordR) = True Then
3040            GDBR.Open "DSN=" + sServerR + ";UID=" + sLoginR + ";PWD=" + _
          sPasswordR
                 List1.AddItem "remote ”’Õƒ≈”« œ "
3050            Exit Do
3060      MsgBox "Connection Successful", vbInformation, "ODBC Logon"
3070      Else

                  Me.Caption = "–—œ”–¡»≈…¡ ”’Õƒ≈”«” Ã≈ REMOTE SERVER " + Format(Now, "HH:MM")
                  DoEvents
          '      MsgBox "Connection Failed", vbExclamation, "ODBC Logon"
3080      End If
           DoEvents
3090  Loop

3100  Me.Caption = "¡–œÃ¡ —’”Ã≈Õœ” œ -‘≈Àœ” ”’Õƒ≈”≈ŸÕ"





      '120   GDBR.Open g2CONNECT

3110   F_RUNNING = 0

3120  Exit Sub

topikos:
'Unload Me
GoTo top 'run_again
Exit Sub


' 3130  MsgBox Str(Erl) + "-----" + Err.Description
'3140  End

remote:

'Unload Me

'run_again

GoTo top
Exit Sub

'3150  MsgBox Str(Erl) + "-----" + Err.Description
'3160  End
       

End Sub



Sub run_again()
   List1.Clear
   List1.BackColor = vbYellow
   open_data
   
End Sub

Private Sub PONTOIthes_Click()

'
'
'
'
'
'
'Dim R As New ADODB.Recordset
'Dim RR As New ADODB.Recordset
'Dim K As Long
'Dim L As Long
'
'Dim Reid As New ADODB.Recordset
'Dim eidh As String
'
'If SERV1 > 0 Then
'
'
'
'   R.Open "SELECT * FROM DIATAKT WHERE ENHM IS NULL ORDER BY ID", GDB1, adOpenDynamic, adLockOptimistic
'   Dim S As Long
'  Do While Not R.EOF
'
'   'Ã¡∆≈’Ÿ ‘œ’”  Ÿƒ… œ’” –œ’ ¡√œ—¡”≈
'   eidh = ""
'   Reid.Open "select DISTINCT KODE from EGGTIM WHERE ATIM='" + R!ATIM + "' AND TIMM>0", GDB1, adOpenDynamic, adLockOptimistic
'   Do While Not Reid.EOF
'      eidh = eidh + Reid!KODE + ";"
'      Reid.MoveNext
'   Loop
'   Reid.Close
'
'
'    GDBR.Execute "INSERT INTO DIATAKT (MHXANH,PELATHS,PONTOI,HME,HISTORY ) VALUES (41,'" + R("PELATHS") + "'," + Str(R("PONTOI")) + ",'" + Format(R("HME"), "MM/DD/YYYY") + "','" + eidh + "')", L
'    If L > 0 Then
'       GDBR.Execute "UPDATE EID SET PONTOI=PONTOI+" + Str(R("PONTOI")) + " WHERE KOD='" + R("PELATHS") + "'", K
'       If K > 0 Then
'          GDB1.Execute "UPDATE DIATAKT SET UPD=1 WHERE ID=" + Str(R("ID"))
'          S = S + R("PONTOI")
'          DoEvents
'       End If
'    End If
'    R.MoveNext
'  Loop
'   List1.AddItem "1. –œÕ‘œ… " + Str(S)
'      List1.AddItem "=================================================="
'   R.Close
'
'
'End If
'
'
'
'If SERV2 > 0 Then
'
'
'
'   R.Open "SELECT * FROM DIATAKT WHERE ENHM IS NULL ORDER BY ID", GDB2, adOpenDynamic, adLockOptimistic
'   Dim S As Long
'  Do While Not R.EOF
'
'   'Ã¡∆≈’Ÿ ‘œ’”  Ÿƒ… œ’” –œ’ ¡√œ—¡”≈
'   eidh = ""
'   Reid.Open "select DISTINCT KODE from EGGTIM WHERE ATIM='" + R!ATIM + "' AND TIMM>0", GDB2, adOpenDynamic, adLockOptimistic
'   Do While Not Reid.EOF
'      eidh = eidh + Reid!KODE + ";"
'      Reid.MoveNext
'   Loop
'   Reid.Close
'
'
'    GDBR.Execute "INSERT INTO DIATAKT (MHXANH,PELATHS,PONTOI,HME,HISTORY ) VALUES (42,'" + R("PELATHS") + "'," + Str(R("PONTOI")) + ",'" + Format(R("HME"), "MM/DD/YYYY") + "','" + eidh + "')", L
'    If L > 0 Then
'       GDBR.Execute "UPDATE EID SET PONTOI=PONTOI+" + Str(R("PONTOI")) + " WHERE KOD='" + R("PELATHS") + "'", K
'       If K > 0 Then
'          GDB2.Execute "UPDATE DIATAKT SET UPD=1 WHERE ID=" + Str(R("ID"))
'          S = S + R("PONTOI")
'          DoEvents
'       End If
'    End If
'    R.MoveNext
'  Loop
'   List1.AddItem "2. –œÕ‘œ… " + Str(S)
'   R.Close
'   List1.AddItem "=================================================="
'
'End If
'
'
'
'
'
'
'If SERV3 > 0 Then
'
'
'
'   R.Open "SELECT * FROM DIATAKT WHERE ENHM IS NULL ORDER BY ID", GDB3, adOpenDynamic, adLockOptimistic
'   Dim S As Long
'  Do While Not R.EOF
'
'   'Ã¡∆≈’Ÿ ‘œ’”  Ÿƒ… œ’” –œ’ ¡√œ—¡”≈
'   eidh = ""
'   Reid.Open "select DISTINCT KODE from EGGTIM WHERE ATIM='" + R!ATIM + "' AND TIMM>0", GDB3, adOpenDynamic, adLockOptimistic
'   Do While Not Reid.EOF
'      eidh = eidh + Reid!KODE + ";"
'      Reid.MoveNext
'   Loop
'   Reid.Close
'
'
'    GDBR.Execute "INSERT INTO DIATAKT (MHXANH,PELATHS,PONTOI,HME,HISTORY ) VALUES (43,'" + R("PELATHS") + "'," + Str(R("PONTOI")) + ",'" + Format(R("HME"), "MM/DD/YYYY") + "','" + eidh + "')", L
'    If L > 0 Then
'       GDBR.Execute "UPDATE EID SET PONTOI=PONTOI+" + Str(R("PONTOI")) + " WHERE KOD='" + R("PELATHS") + "'", K
'       If K > 0 Then
'          GDB3.Execute "UPDATE DIATAKT SET UPD=1 WHERE ID=" + Str(R("ID"))
'          S = S + R("PONTOI")
'          DoEvents
'       End If
'    End If
'    R.MoveNext
'  Loop
'   List1.AddItem "3. –œÕ‘œ… " + Str(S)
'   R.Close
'   List1.AddItem "=================================================="
'
'End If
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'












End Sub











Private Sub PONTOI_Click()

On Error GoTo SEELINE

Dim R As New ADODB.Recordset
Dim RR As New ADODB.Recordset
Dim K As Long
Dim L As Long
List1.AddItem "ÂÎÂ„˜ÔÚ ¸ÌÙ˘Ì"
R.Open "SELECT * FROM DIATAKT WHERE PELATHS LIKE '91%' AND UPD IS NULL ORDER BY ID", GDB, adOpenDynamic, adLockOptimistic
Dim S As Long
Dim nc As Long
nc = 1
Dim mmid As Long

Dim EIDH As String

Dim REID As New ADODB.Recordset
Dim MATIM As String
Dim RATIM As String

Dim RATIM2 As String


Do While Not R.EOF
 
 
' ' SELECT  * FROM EGGTIM WHERE KODE='9139393234844' AND ATIM='L0031'
   
   RATIM = Trim(R!ATIM)
'¬—…” Ÿ ‘«Õ ¡–œƒ≈…Œ« ‘œ’ –≈À¡‘« GIA NA BRO THN –—¡√Ã¡‘… « ¡–œƒ≈…Œ« ATIM2=AR.POS

    REID.Open "select * from EGGTIM WHERE KODE='" + Trim(R!PELATHS) + "' AND HME='" + Format(R!HME, "MM/DD/YYYY") + "' AND ATIM = '" + RATIM + "' ", GDB, adOpenDynamic, adLockOptimistic
    RATIM2 = REID!ATIM2
    REID.Close
   
   
   

'  'BRISKV THN ¡–œƒ≈…Œ« ‘œ’ –≈À¡‘«
'   REID.Open "select * from EGGTIM WHERE KODE='"+R!PELATHS+"' AND HME='"+ATIM='" + R!ATIM + "' AND TIMM>0", GDB1, adOpenDynamic, adLockOptimistic
'
'
  'Ã¡∆≈’Ÿ ‘œ’”  Ÿƒ… œ’” –œ’ ¡√œ—¡”≈
   EIDH = ""
   REID.Open "select DISTINCT KODE from EGGTIM WHERE HME='" + Format(R!HME, "MM/DD/YYYY") + "' AND ATIM='" + RATIM + "' AND ATIM2='" + RATIM2 + "'  AND TIMM>0 ", GDB, adOpenDynamic, adLockOptimistic
   Do While Not REID.EOF
      EIDH = EIDH + REID!KODE + ";"
      REID.MoveNext
   Loop
   REID.Close
'


List1.AddItem "ÂÎÂ„˜ÔÚ ¸ÌÙ˘Ì" + Str(nc)

GDBR.Execute "INSERT INTO DIATAKT (MHXANH,KOLA,IDYPOK,PELATHS,PONTOI,HME,HISTORY) VALUES (" + F_MHXANH + "," + F_APOT + "," + Str(R!ID) + ", '" + R("PELATHS") + "'," + Str(R("PONTOI")) + ",'" + Format(R("HME"), "MM/DD/YYYY") + "','" + EIDH + "')", L
If L > 0 Then
   GDBR.Execute "UPDATE EID SET PONTOI=PONTOI+" + Str(R("PONTOI")) + " WHERE KOD='" + R("PELATHS") + "'", K
   If K > 0 Then
      GDB.Execute "UPDATE DIATAKT SET UPD=1 WHERE ID=" + Str(R("ID"))
      S = S + R("PONTOI")
      DoEvents
   End If
End If

mmid = R!ID
R.MoveNext

If R.EOF Then
     Exit Do
Else
     
     '·Ì ‰ÂÌ ÂÒ·ÙÁÛÂ ÙÔ id ÙÔÙÂ ‚„ÂÚ ·Ô ÙÔ loop
     If mmid = R!ID Then
        Exit Do
     End If
End If





nc = nc + 1

     

Loop

List1.AddItem "–œÕ‘œ… " + Str(S)
R.Close

Exit Sub

SEELINE:

     On Error Resume Next

     GDB.Close
     GDBR.Close
     
     Me.Caption = Str(Erl) + "-----" + Err.Description
     SAVE_ERROR Err.Description & " in PONTOI_Click " & "at line " & Erl
     
     open_data
     




End Sub

Private Sub Timer1_Timer()
'3170     On Error GoTo 0
'3180   If F_RUNNING = 0 And EID2.Value = vbUnchecked Then
'


If F_PONTOI = 1 Then
            Me.Caption = "≈Õ«Ã≈—Ÿ”« –œÕ‘ŸÕ  " + Format(Now, "HH:MM")
            PONTOI_Click
            DoEvents
End If


If F_SYNCHRO_KIN = 1 Then
            Me.Caption = "≈Õ«Ã≈—Ÿ”«  …Õ«”≈ŸÕ  " + Format(Now, "HH:MM")
3190        synchro
            DoEvents
'           Me.Caption = "≈Õ«Ã≈—Ÿ”« ƒ…¡√—¡ÃÃ≈ÕŸÕ  …Õ«”≈ŸÕ  " + Format(Now, "HH:MM")
'3200       DELEGGTIM
'           DoEvents
End If


If F_ENHM_EIDH = 1 Then
           Me.Caption = "≈Õ«Ã≈—Ÿ”« BARCODES  " + Format(Now, "HH:MM")
3210      EnhmBarcodes
           DoEvents
           Me.Caption = "≈Õ«Ã≈—Ÿ”« EIƒŸÕ  " + Format(Now, "HH:MM")
3220      EID_Click
           DoEvents
3230       List1.AddItem "---------------------------------"
End If


If F_LOADPEL = 1 Then


  'ÏÈ· ˆÔÒ· ÙÁÌ ˘Ò· ÛÙ›ÎÌÂÈ ÙÈÚ ÍÈÌÁÛÂÈÚ
        If Val(Right(Format(Now, "HH:MM"), 2)) <= 2 Then
             List1.AddItem "¡–œ”‘œÀ« ‘…ÃŸÕ ”≈ POS"
             LOAD_PELATES
             TIMES_UPD
      
        End If


End If







If F_POS = 1 Then
    Command2_Click  ' SYLLOGH APO IPOS
End If



'3240   End If
End Sub

Private Function IsValidODBCLogin(ByVal sDSN As String, ByVal sUID As String, _
    ByVal sPWD As String) As Boolean
            Dim henv As Long    'Environment Handle
            Dim hdbc As Long    'Connection Handle
            Dim iResult As Integer

              'Obtain Environment Handle
3250          iResult = SQLAllocEnv(henv)
3260          If iResult <> SQL_SUCCESS Then
3270            IsValidODBCLogin = False
3280            Exit Function
3290          End If

              'Obtain Connection Handle
3300          iResult = SQLAllocConnect(henv, hdbc)
3310          If iResult <> SQL_SUCCESS Then
3320            IsValidODBCLogin = False
3330            iResult = SQLFreeEnv(henv)
3340            Exit Function
3350          End If

              'Test Connect Parameters
3360          iResult = SQLConnect(hdbc, sDSN, Len(sDSN), sUID, Len(sUID), sPWD, _
          Len(sPWD))
3370          If iResult <> SQL_SUCCESS Then
3380            If iResult = SQL_SUCCESS_WITH_INFO Then
                  'The Connection has been successful, but SQLState Information
                  'has been returned
                  'Obtain all the SQLState Information
                  'If Check.Value Then ShowSQLErrorInfo hdbc, vbInformation
3390               ShowSQLErrorInfo hdbc, vbInformation
3400              IsValidODBCLogin = True
3410            Else
                  'Obtain all the Error Information
                  'If Check.Value Then ShowSQLErrorInfo hdbc, vbExclamation
3420              ShowSQLErrorInfo hdbc, vbExclamation
3430              IsValidODBCLogin = False
3440            End If
3450          Else
3460            IsValidODBCLogin = True
3470          End If

              'Free Connection Handle and Environment Handle
3480          iResult = SQLFreeConnect(hdbc)
3490          iResult = SQLFreeEnv(henv)

      End Function

'      Private Sub Form_Load()
'
'        Text1.Text = "DSN"
'        Text2.Text = "User ID"
'        Text3.Text = ""
'        Text3.PasswordChar = "*"
'        Command1.Caption = "Test Connect"
'        Check.Caption = "Return Errors and Warnings"
'
'      End Sub

'      Private Sub Command1_Click()
'      Dim sServer As String, sLogin As String, sPassword As String
'
'        sServer = Text1.Text
'        sLogin = Text2.Text
'        sPassword = Text3.Text
'
'        If IsValidODBCLogin(sServer, sLogin, sPassword) = True Then
'          MsgBox "Connection Successful", vbInformation, "ODBC Logon"
'        Else
'          MsgBox "Connection Failed", vbExclamation, "ODBC Logon"
'        End If
'
'      End Sub

      Private Sub ShowSQLErrorInfo(hdbc As Long, iMSGIcon As Integer)
            Dim iResult As Integer
            Dim hstmt As Long
            Dim sBuffer1 As String * 16, sBuffer2 As String * 255
            Dim lNative As Long, iOutlen As Integer

3500          sBuffer1 = String$(16, 0)
3510          sBuffer2 = String$(256, 0)

3520          Do 'Cycle though all the Errors
3530            DoEvents
3540            Me.Refresh

3550            If Check.Value = vbChecked Then
3560               End
3570            End If
3580            iResult = SQLError(0, hdbc, hstmt, sBuffer1, lNative, sBuffer2, _
          256, iOutlen)
3590            If iResult = SQL_SUCCESS Then
3600              If iOutlen = 0 Then
3610                MsgBox "Error -- No error information available", iMSGIcon, _
          "ODBC Logon"
3620              Else
                    ' MsgBox Left$(sBuffer2, iOutlen), iMSGIcon, "ODBC Logon"
                    ' DoEvents
3630                MILSEC 1000
3640                List1.AddItem "·‰ıÌ·ÙÁ ÛıÌ‰ÂÛÁ ÏÂ ·ÔÏ·ÍÒıÛÏÂÌÔ"
3650              End If
3660            End If
3670          Loop Until iResult <> SQL_SUCCESS

      End Sub


Sub HandleError(ByVal proc As String)
3680      On Error Resume Next

3690      Open "c:\MERCVB\errsyn.txt" For Append As #112

3700      Print #112, proc + "-" + Format(Now, "dd/mm/yyyy  hh:mm") + "-" + _
              Format(Erl, "0000") + "-" + Err.Description
3710      Close #112

End Sub


'
'Update [2005].[dbo].[EID]
'   SET [KOD] = <KOD, nvarchar(14),>
'                  ,[KODSYNOD] = <KODSYNOD, nvarchar(14),>
'      ,[HPAR] = <HPAR, smalldatetime,>



Private Sub UPDSTR(ByVal CC As String, ByVal C2, ByVal KOD As String)
      Dim f
3720      If Not IsNull(C2) Then
             C2 = Replace(C2, "'", "~")


                 
           'If fRL(CC) <> C2 Then
           
           'If CC = "ONO" Then
3730         GDB.Execute "UPDATE EID SET " + CC + "='" + C2 + "' where KOD='" + Trim(KOD) + _
                 "';", f
          ' End If
                 
             'F_SQL = F_SQL + CC + "='" + C2 + "',"
                 
3740      End If


End Sub


Private Sub UPDNUM(ByVal CC As String, ByVal C2, ByVal KOD As String)
      Dim f
3750      If Not IsNull(C2) Then
           '  if c2
           'If fRL(CC) <> C2 Then
3760           GDB.Execute "UPDATE EID SET " + CC + "=" + Str(C2) + " where KOD='" + _
                 KOD + "';", f
           'End If
               ' F_SQL = F_SQL + CC + "=" + Str(C2) + ","
3770      End If


End Sub

Private Sub UPDDAT(ByVal CC As String, ByVal C2, ByVal KOD As String)
      Dim f
3780      If Not IsNull(C2) Then
            'If fRL(CC) <> C2 Then
3790           GDB.Execute "UPDATE EID SET " + CC + "='" + Format(C2, "MM/DD/YYYY") + _
                 "' where KOD='" + KOD + "';", f
            'End If
            ' F_SQL = F_SQL + CC + "='" + Format(C2, "MM/DD/YYYY") + ","
3800      End If


End Sub




Function CNULL(ByVal C As String)
If IsNull(C) Then C = ""
CNULL = C

End Function


Sub LOAD_PELATES()
        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    –≈—…√—¡÷«                                                              G
'OK----03       087         08                                    ‘…Ã«                                                                D,###.00
'OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
'OK----17       095         01                                    ‘Ã«Ã¡                                                              I
'----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
'----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
'----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
'OK  ----25       00115                                    ¬  Ÿƒ… œ”


On Error GoTo SEELINE

Dim R22 As New ADODB.Recordset


            Open "C:\MERCVB\points.upd" For Output As #2

            Open "C:\MERCVB\customer.upd" For Output As #1
             ' ena eidos
              R22.Open "SELECT  * FROM EID  where  KOD LIKE '913%'  ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                'On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim mkod As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String

Dim cp As String
Dim nc As Long




            Do While Not R22.EOF
                CC = Space(250)
                cp = Space(89)
                
                
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22("kod")) Then
                mkod = "."
             Else
                mkod = R22("kod")
             End If
                
                
'              If mkod = "005.429" Then
'                 nc = nc + 1
'              End If
                
             mkod = Replace(mkod, Chr(10), " ")
             mkod = Replace(mkod, Chr(13), " ")
                
                
              Mid(CC, 1, 1) = "0"  'Left(mkod + Space(15), 15)
              '  Mid(CC, 3, 15) = Left(mkod + Space(16), 15)   '      Left(mID(mkod, 8, 6) + Space(15), 15)
              Mid(CC, 3, 15) = Left(mkod + Space(15), 15)
              
              Mid(CC, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
              
              
              
              Mid(CC, 246, 2) = "00"
              Mid(CC, 244, 1) = "1"
              Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
              
              
              
              
              Mid(cp, 1, 1) = "0"  'Left(mkod + Space(15), 15)
              Mid(cp, 3, 15) = Left(mkod + Space(15), 15)
              Mid(cp, 19, 20) = Left("–≈À¡‘«” À…¡Õ… «”" + Space(15), 20)
              
              Mid(cp, 60, 1) = "0"
              
              If IsNull(R22("pontoi")) Then
              
                  Mid(cp, 66, 5) = "0"
              Else
                  Mid(cp, 66, 5) = Right("      " + Str(R22("pontoi")), 5)
              
              End If
              
              'Mid(cp, 66, 5) = "0"
              
              
              'Mid(CC, 201, 16) = Left(mkod + Space(16), 16)
              
              
              
              
              
              
 '             0=0   3->15  kodikos  customer.upd   data
'60=>0      66 =>5


              
              
              
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             
             
            ' If Len(CC) < 140 Then
                 Print #2, cp
             'Else
                 Print #1, CC
             'End If

                  
                R22.MoveNext
                
                
                ' If nc > 10 Then Exit Do
                
            Loop






             Close #1
             
             Close #2
             
             
             
             R22.Close
             


If FolderExists("\\POS1\TEC_POS\DATA") Then
     FileCopy "C:\MERCVB\customer.upd", "\\POS1\TEC_POS\DATA\customer.upd"
      FileCopy "C:\MERCVB\points.upd", "\\POS1\TEC_POS\DATA\points.upd"
     List1.AddItem "≈Õ«Ã. ‘œ POS1 ME –≈À¡‘≈”"
Else
     List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS1 ME –≈À¡‘≈”"
End If

If FolderExists("\\POS1\TEC_POS\DATA") Then
      FileCopy "C:\MERCVB\customer.upd", "\\POS2\TEC_POS\DATA\customer.upd"
      FileCopy "C:\MERCVB\points.upd", "\\POS2\TEC_POS\DATA\points.upd"
      List1.AddItem "≈Õ«Ã. ‘œ POS2 ME –≈À¡‘≈”"
Else
      List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS2 ME –≈À¡‘≈”"
End If

If FolderExists("\\POS3\TEC_POS\DATA") Then
      FileCopy "C:\MERCVB\customer.upd", "\\POS3\TEC_POS\DATA\customer.upd"
      FileCopy "C:\MERCVB\points.upd", "\\POS3\TEC_POS\DATA\points.upd"
      List1.AddItem "≈Õ«Ã. ‘œ POS3 ME –≈À¡‘≈”"
Else
      List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS3 ME –≈À¡‘≈”"
End If




                
                Exit Sub
                
                
                
SEELINE:
430     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"

440     On Error Resume Next

SAVE_ERROR Err.Description & " in LOAD_PELATES " & "at line " & Erl

450     GDB.Close
460     GDBR.Close
470     List1.AddItem Str(Erl) + "-----" + Err.Description
480     open_data
               
End Sub


'Private Sub CheckFolders()
'  If Not FolderExists("\\s-fk-fin-ti\private\" & TextBox1.Text) Then
'    MsgBox ("\\s-fk-fin-ti\private\" & TextBox1.Text)
'  End If
'
'  If Not FolderExists("\\s-fk-fin-ti\private\" & TextBox2.Text) Then
'    MsgBox ("\\s-fk-fin-ti\private\" & TextBox2.Text)
'  End If
'
'  If Not FolderExists("\\s-fk-fin-ti\private\" & TextBox3.Text) Then
'    MsgBox ("\\s-fk-fin-ti\private\" & TextBox3.Text)
'  End If
'End Sub

Private Function FolderExists(sFullPath As String) As Boolean
Dim myFSO As Object
Set myFSO = CreateObject("Scripting.FileSystemObject")
FolderExists = myFSO.FolderExists(sFullPath)
End Function







Sub TIMES_UPD()
        ' If Len(Dir("C:\MERCVB\TEST.EXE", vbNormal)) > 0 And cXRONIES.Value = vbChecked Then

'OK ----00 START=021     LEN=13    FIX=                            BARCODE                                                                 A
'OK----01       036         30                                    –≈—…√—¡÷«                                                              G
'OK----03       087         08                                    ‘…Ã«                                                                D,###.00
'OK ----10       095         01                                    ÷–¡ 1=6,2=13,3=24,5=0                                                                A
'OK'' ----12       107         02                                    ÃœÕ¡ƒ¡ Ã≈‘—«”«” 1=‘≈Ã 2= …À¡                                                    A
'OK----17       095         01                                    ‘Ã«Ã¡                                                              I
'----20       140         01                                    ≈À≈’»≈—« ‘…Ã« 0=œ◊… 1=Õ¡…                                                        I,S0~1
'----21       137         01        0                           ∆’√…∆œÃ≈Õœ ≈…ƒœ” 0=œ◊… 1=Õ¡…                                                              I
'----24       131         04        0                           –œÕ‘œ… ≈…ƒœ’”                                                           L
'OK  ----25       00115                                    ¬  Ÿƒ… œ”


On Error GoTo SEELINE


Dim R22 As New ADODB.Recordset


            Open "C:\MERCVB\ERRORSPOSFILE.TXT" For Output As #2

            Open "C:\MERCVB\POSFILE.TXT" For Output As #1
             ' ena eidos
              R22.Open "SELECT  EID.KOD,BARCODES.ERG,ONO,LTI5,( CASE WHEN MON IS NULL THEN  '1' ELSE '1' END )  AS MON,( CASE WHEN FPA=4 THEN 1 ELSE ( CASE WHEN FPA=1 THEN 2 ELSE 3 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID RIGHT JOIN BARCODES ON EID.KOD=BARCODES.KOD where  NOT ( EID.KOD LIKE '913%' ) ORDER BY EID.KOD ", GDB, adOpenDynamic, adLockOptimistic  ' WHERE EID.KOD='" + Text1(0) + "'"
             
             ' ola ta eidh
             'R22.Open "SELECT KOD,ONO,LTI5,( CASE WHEN MON IS NULL THEN  'TEM' ELSE MON END )  AS MON,( CASE WHEN FPA=1 THEN 13 ELSE ( CASE WHEN FPA=5 THEN 0 ELSE 23 END ) END ) AS FPA1 , ( CASE WHEN FPA=5 THEN 5 ELSE  FPA+1 END)  AS TMHMA FROM EID  ", Gdb, adOpenDynamic, adLockOptimistic
                'On Error GoTo iliadis_error


Dim mLTI5 As Single
Dim MONO As String
Dim mkod As String
Dim mbar As String
Dim MFPA As Integer
Dim MBARC As String


Dim nc As Long


            Do While Not R22.EOF
                ' Open "C:\POSFILE.TXT" For Output As #1
                
             If IsNull(R22(0)) Then
                mkod = "."
             Else
                mkod = R22(0)
             End If
                
                
              If mkod = "005.429" Then
                 nc = nc + 1
              End If
                
             mkod = Replace(mkod, Chr(10), " ")
             mkod = Replace(mkod, Chr(13), " ")
                
                
             CC = Left(mkod + Space(21), 20)  ' KVDIKOS
             
             
             
             'R22 ("ONO")
             If IsNull(R22("ono")) Then
                MONO = "."
             Else
                MONO = R22("ono")
             End If
             MONO = Replace(MONO, Chr(10), " ")
             MONO = Replace(MONO, Chr(13), " ")
             
             'R22 ("ONO")
             If IsNull(R22(1)) Then
                MBARC = "0000"
             Else
                MBARC = R22(1)
             End If
             
             MBARC = Replace(MBARC, Chr(10), " ")
             MBARC = Replace(MBARC, Chr(13), " ")

             
             
             CC = CC + Left(MBARC + Space(21), 15) ' BARCODE
             
             
             
             
             
             CC = CC + Mid(MONO + String(30, " "), 1, 30) + String(21, " ") ' Space(21)
             If IsNull(R22("LTI5")) Then
                mLTI5 = 0
             Else
                mLTI5 = R22("LTI5")
             End If
             
             CC = CC + Replace(Format(mLTI5, "00000.00"), ".", ",") '94   '+ String(F_DEK_LIANIKIS, "0"))
     
                '  If IsNull(R22("mon")) Then
                '     CC = CC + Space(3)
                '   Else
                '      CC = CC + R22("mon") '+ Space(3), 3)
                ' End If
     
             
             
             If IsNull(R22("FPA1")) Then
                MFPA = 3
             Else
                MFPA = R22("FPA1")
             End If
             
             
             DoEvents
             
             Me.Caption = nc
             nc = nc + 1
             
             CC = CC + Format(MFPA, "0") + Space(10)
             CC = CC + " 1"    ' monada Format(R22("mon")
             
             
   '          On Error Resume Next
             
            ' CC = CC + Space(18) + Format(R22("tmhma"), "0.00")
             
             CC = CC + "                       00    0  0      "   ' ÔÌÙÔÈ Êı„ÈÊÔÏÂÌ·
             CC = Replace(CC, Chr(13), "")
             If Len(CC) < 140 Then
                 Print #2, CC
             Else
                 Print #1, CC
             End If
'123456789012340078900230567890
                  
                R22.MoveNext
            Loop






             Close #1
             
             Close #2
             
             
             
             R22.Close
             
             
             On Error Resume Next
             


If FolderExists("\\POS1\TEC_POS\DATA") Then
      FileCopy "C:\MERCVB\customer.upd", "\\POS1\TEC_POS\DATA\POSFILE.TXT"
      List1.AddItem "≈Õ«Ã. ‘œ POS1 ME ≈…ƒ«"
Else
      List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS1 ME ≈…ƒ«"
End If

If FolderExists("\\POS2\TEC_POS\DATA") Then
      FileCopy "C:\MERCVB\customer.upd", "\\POS2\TEC_POS\DATA\POSFILE.TXT"
      List1.AddItem "≈Õ«Ã. ‘œ POS2 ME ≈…ƒ«"
Else
      List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS2 ME ≈…ƒ«"
End If

If FolderExists("\\POS3\TEC_POS\DATA") Then
      FileCopy "C:\MERCVB\customer.upd", "\\POS3\TEC_POS\DATA\POSFILE.TXT"
      List1.AddItem "≈Õ«Ã. ‘œ POS3 ME ≈…ƒ«"
Else
      List1.AddItem "!!! ƒEN ≈Õ«Ã. ‘œ POS3 ME ≈…ƒ«"
End If
             
             
             
             
420     Exit Sub

SEELINE:
        'HandleError "MdiForm-load"
        'Resume Next
430     List1.AddItem Format(Now, "dd/mm/yyyy HH:MM") + "·ÔÛıÌ‰ÂÛÁ"

440     On Error Resume Next

        SAVE_ERROR Err.Description & " in TIMES_UPD " & "at line " & Erl
450     GDB.Close
460     GDBR.Close
470     List1.AddItem Str(Erl) + "-----" + Err.Description
480     open_data

490     Exit Sub    'Resume Next








                On Error Resume Next
                
                Exit Sub

         '   End If

End Sub
