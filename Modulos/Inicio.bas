Attribute VB_Name = "Coneccion"
Global LoginSucceeded As Boolean
Global wuser As String
Global wclave As String
Global wmoncont As String
Global wcia As String
Global wruc As String
Global wGrupoPla As String
Global wbus As String
Global wobra As Boolean
Global wCodSystem As String * 2
Global gsFileINI                As String
Global Carga_Cia As String
Global gsSQL_Server As String
Global gsSQL_DB As String
Global TitMsg As String
Global wCancelsis As Boolean
Global wTipoPla As String
Global wtipodoc As Boolean
Global NOMBREEMPRESA As String
Global sueldominimo As Currency
Global porcasigfamiliar As Currency
Public Cadena As String
Public Path_CIA As String
Public Path_Excel As String
Public Path_Reports As String
Global wNomBd As String
Global wserver As String
Global WDatabase As String
Global Sql As String
Global wImprimeBat

Public Objcn As New ADODB.Connection
Public rutaRPT$

Global nom_BD As String

Global cn2 As New ADODB.Connection
Global cn As New ADODB.Connection
Global rs As ADODB.Recordset

Global NameForm As String

Global gsAdminDB As String
Global FormatFecha As String
Global FormatFechaSql As String
Global FormatTimei As String
Global FormatTimef As String
Global wInicioTrans As String
Global wFinTrans As String
Global wCancelTrans As String
Global FechaSys As String
Global wFuncdia As String
Global wAdmin As String
Global wPrintFile As String
Global wSenati As String
Global wMaeGen As Boolean
Global wNamePC As String
Global LetraChica As String
Global SaltaPag As String
Global aPrinloc(1 To 10)        As String



'IMPLEMENTACION GALLOS
Global wImportarXls As Boolean
Global wImportaUti As Boolean

Declare Function SetWindowPos Lib "user32" (ByVal h&, ByVal hb&, ByVal x&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal F&) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function CopyFile Lib "kernel32" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Sub Main()
Dim rstTipo As ADODB.Recordset
On Error GoTo SplashLoadErr
wCancelsis = False

'Cargar INI
Call LoadINISettings

    If App.PrevInstance Then
        MsgBox "El Sistema de Planillas ya esta en uso", vbExclamation, "Sistema de Planilla"
        End
    End If

Load Frmacceso
Frmacceso.Show vbModal
  

If (wuser) = "" And LoginSucceeded Then
   MsgBox "Ingrese Un Usuario", vbInformation, "SISTEMA DE PLANILLAS"
   Exit Sub
End If

success% = SetWindowPos(Frmacceso.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flags)
success% = SetWindowPos(Frmacceso.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, Flags)
Unload Frmacceso

If wCancelsis = True Then Exit Sub
'If LoginSucceeded = False Then Exit Sub

    wserver = ReadINI("SQL", "SQL_Server")
    WDatabase = ReadINI("SQL", "SQL_Database")
    gsComputer = ReadINI("NET", "NET_ComputerName")
    gsAdminDB = ReadINI("NET", "NET_Admin_DataBase")
    gsHOST = ReadINI("NET", "NET_Host")
    
Select Case StrConv(gsAdminDB, 1)
       Case Is = "MYSQL"
            gsSQL_Database = "DBMySQL"
            'Coneccion a MySQL *****
            Objcn.ConnectionString = "driver={MySQL};" _
            & "server=" & wserver & ";" _
            & "uid=" & Trim(wuser) & ";" _
            & "pwd=" & Trim("") & ";" _
            & "database=" & gsSQL_Database & ""
            '******
            
            FechaSys = "NOW()"
            FormatFecha = "yyyy-mm-dd"
            FormatTimei = " 00:00:00"
            FormatTimef = " 23:59:59"
            wInicioTrans = "BEGIN"
            wFinTrans = "COMMIT"
            wCancelTrans = "Rollback"
            wFuncdia = "dayofmonth"
            wAdmin = "ROOT"
            
       Case Is = "SQL SERVER"
            'gsSQL_Database = "BDplanillas"
            'gsSQL_Database = "PRUEBA1"
            'Coneccion a SQL SERVER ****
            Objcn.ConnectionString = "PROVIDER=SQLOLEDB;" _
            & "SERVER=" & wserver & ";" _
            & "UID=" & Trim(wuser) & ";" _
            & "PWD=" & Trim(wclave) & ";" _
            & "Database=" & WDatabase & ""
            '****************
            gsSQL_DB = WDatabase
            FechaSys = "GETDATE()"
            FormatFecha = "mm/dd/yyyy"
            FormatFechaSql = "MDY" 'PARA LA CONFIGURACION DEL SQL
            FormatTimei = " 12:00:00 AM"
            FormatTimef = " 11:59:59 PM"
            wInicioTrans = "BEGIN TRANSACTION"
            wFinTrans = "COMMIT TRANSACTION"
            wCancelTrans = "Rollback transaction"
            wFuncdia = "day"
            wAdmin = "SA"
            
            
End Select

Objcn.ConnectionTimeout = 100
Objcn.Open
Set cn = Objcn
cn.CommandTimeout = 0



   
   Sql$ = "SELECT * from USERS where login='" & UCase(wuser) & "'  and status<>'*'"
   If UCase(wuser) = wAdmin Then
   Else
      If (fAbrRst(rs, Sql$)) Then
         If Trim(rs!Password) <> Trim(wclave) And StrConv(gsAdminDB, 1) = "MYSQL" Then
            MsgBox "Clave Incorrecta " & Chr(13) & "Asegurese de Escribir bien la Clave Tomando en Cuenta las Mayusculas y/o Minusculas", vbCritical, "Sistema de Planillas"
          Exit Sub
         End If
      Else
         MsgBox "Usuario no Autorizado", vbCritical, "Sistema de Planillas"
         Exit Sub
      End If
      If rs.State = 1 Then rs.Close
   End If
   
If StrConv(gsAdminDB, 1) = "MYSQL" Then
   Sql$ = "Set AUTOCOMMIT = 0"
   cn.Execute Sql$
End If
  
'MDBBBBBBBBBBBB
fc_Limpiar_MDB
wCodSystem = "04"
LetraChica = Chr(15)
SaltaPag = Chr(12) + Chr(13)

'' Fecha del sistema
Sql$ = "set dateformat " & Coneccion.FormatFechaSql & " "
Sql$ = Sql$ & "SELECT " & FechaSys & " "
Set rstTipo = New ADODB.Recordset
Set rstTipo = cn.Execute(Sql$)
If Format(rstTipo(0), FormatFecha) <> Format(Date, FormatFecha) Then
   Screen.MousePointer = 0
   MsgBox "La Fecha de su PC no coincide con la del Día Actual,Por Favor Cambiela con " & Format(rstTipo(0), "DD/MM/YYYY"), vbInformation
   Exit Sub
End If
If rstTipo.State = 1 Then rstTipo.Close

' Moneda Contable
        Sql$ = "SELECT ctemonedac,senati from cia where cod_cia='" & wcia & "' and status<>'*'"
        If (fAbrRst(rs, Sql$)) Then
           wmoncont = rs!ctemonedac
           If IsNull(rs!senati) Then wSenati = "" Else wSenati = rs!senati
        End If
        If rs.State = 1 Then rs.Close
        
SplashLoadErr:
m = Err.Number
      Select Case m
             Case Is = -2147217843
                   MsgBox "Usuario  y/o Clave de Acceso Incorrectos", vbCritical
                   Exit Sub
              Case Is = -2147467259
                   MsgBox "No hay Coneccion con el Servidor", vbCritical
                   Exit Sub
              Case Is = 0
              Case Else
                    MsgBox Err.Description, vbCritical, "Error " & Err.Number
                    Exit Sub
       End Select

'Unload Frmacceso


On Error GoTo FUNKA

Carga_Cia = "Select cod_cia,razsoc,ruc From cia where status<>'*' Order By cod_cia"
FrmSelCia.Show vbModal

Load MDIplared
MDIplared.Show

Sql$ = "SELECT ctemonedac,senati,ruc from cia where cod_cia='" & wcia & "' and status<>'*'"
If (fAbrRst(rs, Sql$)) Then
   wmoncont = rs!ctemonedac
   If IsNull(rs!senati) Then wSenati = "" Else wSenati = rs!senati
End If
If rs.State = 1 Then rs.Close

'verificar usuario con autorizacion
Sql$ = "select * from users where cod_cia='" & wcia & "' and  status<>'*' and name_user='" & wuser & "' and sistema='" & wCodSystem & "'"
cn.CursorLocation = adUseClient
Set rs = New ADODB.Recordset
Set rs = cn.Execute(Sql$, 64)
If rs.RecordCount <= 0 And StrConv(wuser, 1) <> wAdmin Then
   MsgBox "Usuario sin Autorizacion", vbCritical, "Acceso Al sistema de Bancos"
   Load FrmSelCia
   FrmSelCia.Show
   FrmSelCia.ZOrder 0
   If rs.State = 1 Then rs.Close
   Exit Sub
End If
If rs.State = 1 Then rs.Close

wTipoPla = Tipo_Planilla
'' Tipo de Cambio
'SQL$ = "SELECT compra,contable,factor from tipo_cambio where fec_crea " _
' & " BETWEEN '" & Format(Date, FormatFecha) + FormatTimei & "' AND '" & Format(Date, FormatFecha) + FormatTimef & "' " _
' & " AND cod_cia='" & wcia & "'"
'
'cn.CursorLocation = adUseClient
'Set RS = New ADODB.Recordset
'Set RS = cn.Execute(SQL$, 64)
'If RS.RecordCount > 0 Then
'   mtipo_cambio = RS!compra
'   mtc_contable = RS!contable
'   findmon = True
'Else
'   MsgBox "No hay Tipo de Cambio del Día,Por Favor Ingrese", vbInformation
'   Load Frmtipcamb
'   Frmtipcamb.Show
'   Frmtipcamb.ZOrder 0
'End If

MDIplared.Enabled = True
MDIplared.SetFocus
Unload FrmSelCia
MDIplared.Activa_Menu
wGrupoPla = ""
Set rs = New ADODB.Recordset
Set rs = cn.Execute("Select cteigv,ruc from cia where cod_cia='" & wcia & "' and status<>'*'")
If rs.RecordCount <= 0 Then
   MsgBox "No Existen Datos de Compañia", vbCritical, "Planillas"
   Load Frmcia
   Frmcia.Show
   Frmcia.ZOrder 0
Else
   wigv = rs!cteigv * 100
   wruc = Trim(rs!RUC)
End If
Set rs = cn.Execute("Select grupo from plagrupos where ruc='" & wruc & "'")
If rs.RecordCount > 0 Then wGrupoPla = Trim(rs!GRUPO & "")
rs.Close: Set rs = Nothing
Exit Sub
FUNKA:
MsgBox "Error : " & Err.Description, vbCritical, "Planillas"
End Sub
Sub LoadINISettings()
  Dim x As String
    '==================================================
    gsFileINI$ = App.Path & "\Logs\Plared00.INI"
    WgsFileINI$ = "C:\Windows\win.INI"
    wImprimeBat = "N"
    Dim iBat As String
    
    iBat = ReadINI("PRINTERG", "PRINT_BAT")
    If iBat = "S" Then wImprimeBat = "S"
    
  ' NET
    wuser = ReadINI("NET", "NET_UserName")
    wclave = ReadINI("NET", "NET_UserClave")
    gsComputer = ReadINI("NET", "NET_ComputerName")
    
    gsAdminDB = ReadINI("NET", "NET_Admin_DataBase")
    gsHOST = ReadINI("NET", "NET_Host")
    
  'GLOBAL
    wcia = ReadINI("GLOBAL", "GLB_Cia")
    wAlmacen = ReadINI("GLOBAL", "GLB_Almacen")
    wCaja = ReadINI("GLOBAL", "GLB_Caja")
    wclavepago = ReadINI("GLOBAL", "GLB_Clave")
    Mctagana = ReadINI("GLOBAL", "GLB_ctagana")
    Mctaperd = ReadINI("GLOBAL", "GLB_ctaperd")
    wPrintVouvherCheque = UCase(ReadINI("GLOBAL", "GLB_VoucherCheque"))
    
  'PRINTERG
    wPrintCh = Trim(ReadINI("PRINTERG", "GLB_ch"))
    wPrintAs = Trim(ReadINI("PRINTERG", "GLB_as"))
    wPrintVa = Trim(ReadINI("PRINTERG", "GLB_va"))
    wPrintRe = Trim(ReadINI("PRINTERG", "GLB_re"))

  ' SQL
    wserver = ReadINI("SQL", "SQL_Server")
    gsSQL_Database = ReadINI("SQL", "SQL_Database")
    gsSQL_Server = ReadINI("SQL", "SQL_Server")
    gsSQL_Login = ReadINI("SQL", "SQL_Login")
    gsSQL_Password = ReadINI("SQL", "SQL_Password")
    giSQL_LoginTimeOut = ReadINI("SQL", "SQL_LoginTimeOut")
    giSQL_QueryTimeOut = ReadINI("SQL", "SQL_QueryTimeOut")
    giSQL_ConnectionMode = ReadINI("SQL", "SQL_ConnectionMode")
    
  ' MDB y RPT
    rutaRPT$ = ReadINI("MDB / RPT", "RPT_Path")
    nom_BD = App.Path & "\DATA\TEMP.mdb"
    
    'ACCESS
    Path_CIA = ReadINI("ACCESS", "PATH")
    
    'EXCEL GUARDAR
    Path_Excel = ReadINI("EXCEL", "PATH")
  
    'Reportes
    Path_Reports = ReadINI("REPORTS", "PATH")
      
End Sub

'------------------------------------------------------------
'esta función devuelve el valor del  archivo INI para las
'variables Globales
'------------------------------------------------------------
Function ReadINI(cSection$, cKeyName$) As String
   'Se le quita un caracter porque el último es fin de linea
    Dim sRet As String
    Dim Longitud%
    Dim Def$
    
    sRet = String(255, " ")
    Longitud = Len(sRet)
    
   Call GetPrivateProfileString(cSection$, ByVal cKeyName, "", ByVal sRet$, ByVal Len(sRet), gsFileINI)
   
   ReadINI = Left(Trim$(sRet), Len(Trim(sRet)) - 1)
    
End Function
Function WriteINI(cSection$, cKeyName$, cNewString$) As Integer
    WriteINI = WritePrivateProfileString(cSection$, ByVal cKeyName$, ByVal cNewString$, gsFileINI)
End Function
Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys Text, wait
   Set WshShell = Nothing
End Sub
