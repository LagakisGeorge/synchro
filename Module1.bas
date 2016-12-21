Attribute VB_Name = "Module1"
Global GDB1 As New ADODB.Connection

Global GDB2 As New ADODB.Connection
Global GDB3 As New ADODB.Connection

Global GDB As New ADODB.Connection
Global g_stop As Integer


Global GDBR As New ADODB.Connection



Global gCONNECT, g2CONNECT
Declare Function GetCurrentTime Lib "kernel32" Alias "GetTickCount" () As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Sub MILSEC(ByVal MILISECONDS As Long)
          Dim start As Long

3810      start = GetCurrentTime()

3820      Do
3830      Loop Until GetCurrentTime() - start > MILISECONDS




End Sub
Sub main()
'
'g_stop = 0
'Do While True
'
'   synchro.Show
'   If g_stop = 1 Then
'      End
'   End If
'
'
'Loop


End Sub




Function FINDPARAMETROI(KATEG As Integer, FORMA As String, _
                        parametros As String, _
                        DEFAULT As String, _
                        SXOLIA As String)

        '<EhHeader>
        On Error GoTo FINDPARAMETROI_Err

        '</EhHeader>
        ' f_autoNumber = Val(FINDPARAMETROI(1, "PELAT2", "F_autoNumber", "0", "Αρίθμηση αυτόματη 00-00-000  =1 Οχι=0"))
        'F_1ST_CHOICE = Val(FindParametroi(1,"PAR2", "F_1ST_CHOICE", "2", "Πρoεπιλεγμένο παραστατικό")) 'posa psifia tha exei h kathe seira
        Dim R   As New ADODB.Recordset

        Dim sql As String

100     If DEFAULT = "DELETE" Then
110         GDB.Execute "DELETE FROM PARAMETROI WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
120         FINDPARAMETROI = 0

            Exit Function

        End If

130     If DEFAULT = "UPDATE" Then
140         GDB.Execute "UPDATE PARAMETROI SET VAR='" + parametros + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
150         FINDPARAMETROI = 0

            Exit Function

        End If

        Dim N As Long

        'SEIR_SELID1

160     sql = "select * from PARAMETROI where FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
    
170     R.Open sql, GDB, adOpenDynamic, adLockBatchOptimistic

        Dim ll As Long

        ll = R("SXOLIA").DefinedSize

180     If R.EOF Then
190         sql = "insert into PARAMETROI (FORMA,VAR,TIMH,SXOLIA) VALUES ('" + FORMA + "','" + parametros + "','" + DEFAULT + "','" + Left(SXOLIA, ll) + "')"

200         GDB.Execute sql, N

210         FINDPARAMETROI = DEFAULT
212     Else

214         If IsNull(R("TIMH")) Then
216             GDB.Execute "UPDATE PARAMETROI SET TIMH='" + DEFAULT + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'"
218             FINDPARAMETROI = DEFAULT
219         Else
220             FINDPARAMETROI = Trim(R("TIMH"))
221         End If

222     End If

        On Error Resume Next

        'Dim n As Long

230     If Left(SXOLIA, 20) <> Left(R("SXOLIA"), 20) Then
240         GDB.Execute "UPDATE PARAMETROI SET SXOLIA='" + Left(SXOLIA, ll) + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If

250     If IsNull(R("SXOLIA")) Then
260         GDB.Execute "UPDATE PARAMETROI SET SXOLIA='" + SXOLIA + "' WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If
        
        
        If IsNull(R("KATEG")) Then
            GDB.Execute "UPDATE PARAMETROI SET KATEG=1 WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If


        If KATEG <> R("KATEG") Then
            GDB.Execute "UPDATE PARAMETROI SET KATEG=" + Str(KATEG) + " WHERE  FORMA='" + FORMA + "' AND VAR='" + parametros + "'", N
        End If

        






270     R.MoveNext

280     If Not R.EOF Then
290         If parametros = Trim(R("VAR")) Then
       
300             GDB.Execute "DELETE FROM PARAMETROI  WHERE  ID=" + Str(R("ID")), N
            End If
        End If

        '<EhFooter>
        Exit Function

FINDPARAMETROI_Err:
        'MsgBox Err.Description & vbCrLf & _
         "in ADOMERCNEW.kentriko.FINDPARAMETROI " & _
         "at line " & Erl, _
         vbExclamation + vbOKOnly, "Application Error"
        'SAVE_ERROR Err.Description & " in ADOMERCNEW.kentriko.FINDPARAMETROI " & "at line " & Erl

        Resume Next

        '</EhFooter>

End Function


Sub SAVE_ERROR(COMMENT)

    On Error Resume Next

    'SAVE_ERROR Err.Description & " in Project1.Form1.cmdCommand2_Click " & " at line " & Erl
    Dim f As Integer

    f = FreeFile
    Open "C:\MERCVB\ERRSYNCHRO.TXT" For Append As #f
    Write #f, Format(Now, "DD/MM/YYYY HH:MM") + COMMENT

    Close #f

End Sub

