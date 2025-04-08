Attribute VB_Name = "Boleta"

Dim Mstatus As String
Dim macus As Currency
Public ArrDsctoCTACTE() As Variant
Public MAXROW As Long
Public macui As Double
Public Function CalculoAportaciones(ByVal wcia As String, ByVal wcodpla As String, ByVal VArea As String, ByVal VTipo As String, ByVal vaño As Integer, ByVal Vmes As Integer, ByVal Mstatus As String) As String
    'Calculo de Aportaciones
    '------------------------------------------------------------------------------------------------------------------
    MqueryCalA = ""
    For I = 1 To 20
        Sql$ = F03(Trim(wcia), Trim(wcodpla), VArea, VTipo, vaño, Vmes, Mstatus, Format(I, "00"))
        If Sql$ <> "" Then
           cn.CursorLocation = adUseClient
           Set rs = New ADODB.Recordset
           Set rs = cn.Execute(Sql$, 64)
           If rs.RecordCount > 0 Then
              rs.MoveFirst
              If IsNull(rs(0)) Or rs(0) = 0 Then
              Else
                 MqueryCalA = MqueryCalA & "a" & Format(I, "00") & " = " & rs(0) & ","
              End If
           End If
           If rs.State = 1 Then rs.Close
        End If
    Next
    CalculoAportaciones = MqueryCalA
    '-------------------------------------------------------------------------------------------------------------------
End Function
Public Function CalculoDeducciones(ByVal wcia As String, ByVal wcodpla As String, ByVal Mstatus As String, ByVal wcodafp As String, _
                                   ByVal vtope As String, ByVal VCargo, ByVal veesalud As String, ByVal vsindicato As String, ByVal vfecjub As String, _
                                   ByVal wtipodoc As Boolean, ByVal VTipo As String, ByVal vaño As Integer, ByVal Vmes As Integer, ByVal VHorasBol As Integer, ByVal VSemana As String, ByVal sn_quinta As Boolean) As String
Dim MqueryCalD As String
    MqueryCalD = ""
    For I = 1 To 20
        'If i = 11 Then Stop
         Sql$ = F02(Trim(wcia), Trim(wcodpla), Trim(wcodafp), vtope, VCargo, IIf(veesalud = "S", "1", "0"), _
                    IIf(vsindicato = "S", "1", "0"), vfecjub, wtipodoc, VTipo, Mstatus, vaño, Vmes, VHorasBol, Trim(VSemana), sn_quinta, Format(I, "00"))
        
        If Sql$ <> "" Then
           If (fAbrRst(rs, Sql$)) Then
              rs.MoveFirst
              If IsNull(rs(0)) Or rs(0) = 0 Then
                MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
              Else
                 If I = 11 Then
                    For J = 1 To 5
                       MqueryCalD = MqueryCalD & "d" & Format(I, "00") & Format(J, "0") & " = " & rs(J - 1) & ","
                    Next J
                 ElseIf I = 13 And wtipodoc = False Then
                    If rs(0) < 0 Then
                        MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                    Else
                        MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) / 2 & ","
                    End If
                 Else
                    If rs(0) < 0 Then
                        MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & 0 & ","
                    Else
                        MqueryCalD = MqueryCalD & "d" & Format(I, "00") & " = " & rs(0) & ","
                    End If
                 End If
              End If
           End If
           If rs.State = 1 Then rs.Close
        End If
    Next I
CalculoDeducciones = MqueryCalD
End Function
Public Function CalculoIngresos(ByVal wcia As String, ByVal wcodpla As String, ByVal Mstatus As String, _
                                ByVal vaño As Integer, ByVal Vmes As Integer, ByVal VTipotrab As String, ByVal VCargo As String, _
                                ByVal vtope As String, ByVal VSemana As String, ByVal VTipo As String, ByVal wtipodoc As Boolean) As String
Dim MqueryI As String
    'Calculo de ingresos
    '-----------------------------------------------------------------------------------
    MqueryI = ""
    For I = 1 To 50
        Select Case VTipo
               Case Is = "01" 'Normal
                    Sql$ = F01(Trim(wcodpla), Format(I, "00"), Mstatus, vaño, Vmes, VTipotrab, "", VCargo, vtope, Val(VSemana))
               Case Is = "02" 'Vacaciones
                    Sql$ = V01(Trim(wcia), Trim(wcodpla), vtope, VCargo, vaño, Vmes, Mstatus, VTipotrab, Format(I, "00"))
               Case Is = "03" 'Gratificaciones
                    Sql$ = G01(Trim(wcia), Trim(wcodpla), vtope, VCargo, VTipotrab, vaño, Vmes, Mstatus, Format(I, "00"))
            Case Is = "04"
                Sql$ = F01(Trim(wcodpla), Format(I, "00"), Mstatus, vaño, Vmes, VTipotrab, "", VCargo, vtope, Val(VSemana))
            Case Is = "05"
                Sql$ = F01(Trim(wcodpla), Format(I, "00"), Mstatus, vaño, Vmes, VTipotrab, "", VCargo, vtope, Val(VSemana))
        End Select
        
        If Trim(Sql$) <> "" Then
           cn.CursorLocation = adUseClient
           Set rs = New ADODB.Recordset
           Set rs = cn.Execute(Sql$, 64)
           If rs.RecordCount > 0 Then
              rs.MoveFirst
              If IsNull(rs(0)) Or rs(0) = 0 Then
              Else
                If wtipodoc = False And (I = 10 Or I = 11 Or I = 13 Or I = 21 Or I = 22 Or I = 23 Or I = 24) Then
                    MqueryI = MqueryI & "i" & Format(I, "00") & " = " & Round(rs(0) / 2, 2) & ","
                Else
                 MqueryI = MqueryI & "i" & Format(I, "00") & " = " & rs(0) & ","
                End If
               
              End If
           End If
           If rs.State = 1 Then rs.Close
        End If
    Next
    CalculoIngresos = MqueryI
    Exit Function
End Function

Public Function EliminaBoleta(ByVal wtipodoc As Boolean, ByVal Vano As Integer, ByVal Vmes As Integer, ByVal Vdia As Integer, ByVal wcodpla As String, ByVal VPerPago As String, _
            ByVal VTipobol As String, ByVal VSemana As String) As Boolean

Dim Mgrab As String
Dim Sql As String

EliminaBoleta = False
If wtipodoc = True Then

Else
   Sql$ = "select placod from plahistorico " _
   & "where cia='" & wcia & "' and proceso='01' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & wcodpla & "'"
   If (fAbrRst(rs, Sql$)) Then
      MsgBox "Ya se genero la Boleta de Pago " & Chr(13) & "No se Puede Anular el Adelanto de Quincena", vbCritical, "Sistema de Planilla"
      Exit Function
   End If
   Mgrab = MsgBox("Seguro de Eliminar Adelanto de Quincena", vbYesNo + vbQuestion, TitMsg)
End If

Sql$ = wInicioTrans
cn.Execute Sql$
If wtipodoc = True Then
   Select Case VPerPago
          Case Is = "02"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
               & "and status<>'*' and placod='" & wcodpla & "'"
          Case Is = "04"
               Sql$ = "update plahistorico set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
               & "where cia='" & wcia & "' and proceso='" & VTipobol & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
               & "and status<>'*' and placod='" & wcodpla & "'"
   End Select
Else
   Sql$ = "update plaquincena set status='*',fec_modi=" & FechaSys & ",user_modi='" & wuser & "' " _
   & "where cia='" & wcia & "' and Year(fechaproceso) = " & Vano & " And month( fechaproceso) = " & Vmes & " " _
   & "and status<>'*' and placod='" & wcodpla & "'"
End If
cn.Execute Sql$

'If wTipoDoc = True Then
   Sql$ = "select * from plabolcte " _
   & "where cia='" & wcia & "' and placod='" & UCase(Trim(wcodpla)) & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and day( fechaproceso) = " & Vdia & " and status<>'*'"
   If (fAbrRst(rs, Sql$)) Then rs.MoveFirst
   Do While Not rs.EOF
      Sql = "update plactacte set fecha_cancela=null,pago_acuenta=pago_acuenta-" & rs!importe & " where cia='" & wcia & "' and placod='" & UCase(Trim(wcodpla)) & "' and tipo='" & rs!tipo & "' and id_doc=" & rs!id_doc
      cn.Execute Sql$
      rs.MoveNext
   Loop
   Sql$ = "update plabolcte set status='*'" _
   & "where cia='" & wcia & "' and placod='" & UCase(Trim(wcodpla)) & "' and proceso='" & VTipobol & "' and semana='" & VSemana & "' and year(fechaproceso)=" & Vano & " " _
   & "and month( fechaproceso) = " & Vmes & " and day( fechaproceso) = " & Vdia & " and status<>'*'"
   cn.Execute Sql$
'End If


Sql$ = wFinTrans
cn.Execute Sql$
EliminaBoleta = True
End Function
Public Sub devengue(ByVal wcodpla As String, ByVal VFProceso As String)
Dim Sql As String
Dim Mgrab As Integer

Mgrab = MsgBox("Seguro de Grabar Boleta", vbYesNo + vbQuestion, TitMsg)
If Mgrab <> 6 Then Exit Sub
Screen.MousePointer = vbArrowHourglass

Sql$ = wInicioTrans
cn.Execute Sql$

Sql = "select * from plahistorico where cia='" & wcia & "' and proceso='02' and year(fechaproceso)<=" & Val(Mid(VFProceso, 7, 4)) & " and month(fechaproceso)=1 and status='D' and placod='" & wcodpla & "' order by turno"
If (fAbrRst(rs, Sql)) Then rs.MoveFirst
Sql = "update plahistorico set status='T' where cia='" & rs!cia & "' and placod='" & rs!PlaCod & "' and status='" & rs!status & "' and turno='" & rs!turno & "'"
cn.Execute Sql
rs.Close
Sql$ = wFinTrans
cn.Execute Sql$

Screen.MousePointer = vbDefault
End Sub

Public Function F01(ByVal wcodpla As String, concepto As String, _
    ByVal Mstatus As String, ByVal Vano As Integer, ByVal Vmes As Integer, ByVal vtiptrab As String, ByVal VVacacion As String, _
    ByVal VCargo As String, ByVal vtope As String, Optional ByVal pSemana As Integer) As String      'INGRESOS
Dim rsF01 As ADODB.Recordset
Dim mFactor As Currency
Dim nHijos As Integer
Dim RX As New ADODB.Recordset
Dim sSQL As String
Dim pSemanaPago As Integer
mFactor = 0
nHijos = 0
F01 = ""
Select Case concepto
       Case Is = "01" 'BASICO
            F01 = "select round((b.importe/factor_horas)*a.h01,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & Trim(concepto) & "' and b.status<>'*'"
       Case Is = "02" 'ASIGNACION FAMILIAR
            sSQL = "SELECT semana_pago FROM PLACONSTANTE WHERE STATUS!='*' AND CIA='" & wcia & "' and codinterno='02' and tipomovimiento='02'"
            If (fAbrRst(RX, sSQL)) Then pSemanaPago = RX(0)
            If RX.State = 1 Then RX.Close
            
              If Trim(wcodpla) = "SE008" Then Stop
                
            If vtiptrab = "02" Then
                '*******codigo agregado giovanni 15092007***************************************
                'Call Recupera_Tipo_de_Tiempo(wcia, Trim(wcodpla))

                'Select Case Reportes_Centrales.rs_RptCentrales_pub!descrip
                 '   Case "SEMANAL"
                  '      F01 = "select b.importe/(select count(semana) from plasemanas where cia='03'and year(fechaf)=" & Vano & " and month(fechaf)=" & Vmes & ") "
                   '     F01 = F01 & "as basico from platemphist a,plaremunbase b "
                    '    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
                     '   F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                      '  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                    'Case Else
                     '   If Semana_Calcular(pSemana, pSemanaPago, Year(FrmCabezaBol.Cmbfecha), wcia) Then
                      '      F01 = "select round((b.importe/factor_horas)*(factor_horas),2) as basico from platemphist a,plaremunbase b "
                       '     F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
                        '    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                         '   F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                        'End If
                'End Select
                'Set Reportes_Centrales.rs_RptCentrales_pub = Nothing
                '*******************************************************************************
                '/factor_horas)*(factor_horas)
                If Semana_Calcular(pSemana, pSemanaPago, Year(Vano), wcia) Then
                    F01 = "select round((b.importe,2) as basico from platemphist a,plaremunbase b "
                    F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
                    F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                    F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                End If
            Else
                F01 = "select round((b.importe,2) as basico from platemphist a,plaremunbase b "
                F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
                '/factor_horas)*(a.h01+a.h03+a.h08),2
            End If
            
       Case Is = "03" 'ASIGNACION MOVILIDAD
            ' CAMBIO DE H14 POR H01  {>MA<} 10/06/2007
            F01 = "select round((b.importe/factor_horas)*a.H01+A.H03,2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
        
       Case Is = "04" 'BONIFICACION T. SERVICIO
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h02+a.h03+a.h04+a.h05+a.h12),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           
       Case Is = "05" 'INCREMENTO AFP 10.23%

            '*******************codigo Original**********************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case (vtiptrab)
                Case "01": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "02": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '*****************************************************************************************
            
       Case Is = "06" 'INCREMENTO AFP 3%
            '*****************codigo modificado giovanni 29082007**************************************
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            'F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"

            '************codigo agregado giovanni 05082007****************
            Select Case Trim(vtiptrab)
                Case "01": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                Case "02": F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            End Select
            '*************************************************************
            
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03+8),2) as basico from platemphist a,plaremunbase b "
            'F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & Trim(wcia) & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            '******************************************************************************************
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            F01 = "select round((b.importe/factor_horas)*(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
            
       Case Is = "08" 'SOBRETASA (CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
          
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select ROUND(" & mFactor & " *(a.h13),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               'F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h13),2) as basico from platemphist a,plaremunbase b "
               'F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
               'F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               'F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "09" 'DOMINICAL
            If vtiptrab <> "01" Then
               F01 = "select round(((b.importe/factor_horas)*h02),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "10" 'EXTRAS L-S
            
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
''               F01 = "select round((b.importe/factor_horas)* ((" & mFactor & " *(a.h10))+ a.h10),2) as basico from platemphist a,plaremunbase b "
''               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
''               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
''               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               '=====================================
               
                 'bts = convierte_cant
               
'                 F01 = "select round((((b.importe+" & bts & "+A.I07+A.I06)/factor_horas)* " & mFactor & ") *(a.h10),2) as basico,B.IMPORTE,FACTOR_HORAS from platemphist a,plaremunbase b "
'                 F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
'                 F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                 F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.status<>'*'"
                 
                 F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h10,2)) as basico"
                 F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*' AND B.CIA=A.CIA) INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                 F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(wcodpla) & "' AND A.STATUS!='*' and a.concepto not in ('02')"
                 
            End If
       Case Is = "11" 'EXTRAS D-F
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                'bts = convierte_cant
               
'               F01 = "select round((((b.importe+" & bts & ")/factor_horas)* " & mFactor & ") *(a.h11),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
'               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h11,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND B.CIA ='" & wcia & "' AND A.PLACOD='" & Trim(wcodpla) & "' AND A.STATUS!='*' and a.concepto not in ('02')"

            End If
       Case Is = "12" 'FERIADOS
            F01 = "select round((b.importe/factor_horas)*a.h03,2) as basico,A.H03 from platemphist a,plaremunbase b "
            F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & Trim(wcodpla) & "' "
            F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "13" 'REINTEGROS
       Case Is = "14" 'VACACIONES (CONSTRUCCION CIVIL)
            If VVacacion = "S" Then
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "15" 'GRATIFICACION
       Case Is = "29" 'HORARIO DE VERANO
       Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*' AND CIA ='" & wcia & "'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
                
                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h15,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(wcodpla) & "' AND A.STATUS!='*' AND B.CIA='" & wcia & "'"
       End If
       Case Is = "17" 'ASIGNACION ESCOLAR
            If vtiptrab = "01" Then
            End If
            If vtiptrab = "02" Then
            End If
            If vtiptrab = "05" Then
               Sql$ = "select ultima from plasemanas where cia='" & wcia & "' and semana='" & Format(VSemana, "00") & "' and ano='" & Vano & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  If rsF01!ultima = "S" Then
                     Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
                     If (fAbrRst(rsF01, Sql$)) Then
                        nHijos = Numero_Hijos(Trim(wcodpla), "S", "S", VFProceso, 18)
                        mFactor = rsF01!factor
                        If rsF01.State = 1 Then rsF01.Close
                        F01 = "select round((((b.importe/factor_horas)*8)* " & mFactor & ")/12 * " & nHijos & ",2) as basico from platemphist a,plaremunbase b "
                        F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                        F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                        F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
                     End If
                  End If
               End If
            End If
       Case Is = "18" 'UTILIDADES
       Case Is = "20" 'BUC
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
            
        Case Is = "21" '3RO HORAS EXTRAS
                Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
                mFactor = rsF01!factor
                If rsF01.State = 1 Then rsF01.Close
                'select round((((b.importe+" & bts & "+A.I07+A.I06)/factor_horas)* " & mFactor & ") *(a.h10),2) as basico
                  'bts = convierte_cant
                  
'                F01 = "select round(((b.importe+" & bts & ")/factor_horas)* (" & mFactor & " *(a.h17)),2) as basico from platemphist a,plaremunbase b "
'                F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
'                F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'                F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"

                F01 = "SELECT SUM(ROUND(((IMPORTE/(FACTOR_HORAS+CASE WHEN TIPO='02' AND B.tipotrabajador='02' AND a.concepto IN ('05','06') THEN 8 ELSE 0 END))*" & mFactor & ")*c.h17,2)) as basico"
                F01 = F01 & " FROM PLAREMUNBASE A INNER JOIN PLANILLAS B ON (B.PLACOD=A.PLACOD AND B.STATUS!='*') INNER JOIN platemphist C ON (C.PLACOD=A.PLACOD AND B.STATUS!='*' and YEAR(fechaproceso)=" & Vano & " and MONTH(fechaproceso)=" & Vmes & ") "
                F01 = F01 & " WHERE A.CIA='" & wcia & "' AND A.PLACOD='" & Trim(wcodpla) & "' AND A.STATUS!='*' and a.concepto not in ('02')"
                
            End If
        
'       Case Is = "21" 'BONIFICACION POR ALTURA
'            SQL$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(vcargo) & "' and status<>'*'"
'            If (fAbrRst(rsF01, SQL$)) Then
'               mFactor = rsF01!factor
'               If rsF01.State = 1 Then rsF01.Close
'               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h18),2) as basico from platemphist a,plaremunbase b "
'               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
'               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
'               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
'            End If
       Case Is = "22" 'BONIF. CONTACTO AGUA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h19),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "23" 'BONIF.POR ALTITUD
            If VAltitud = "S" Then
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
               If (fAbrRst(rsF01, Sql$)) Then
                  mFactor = rsF01!factor
                  If rsF01.State = 1 Then rsF01.Close
                  F01 = "select round((" & mFactor & " / 8) * (a.h01+a.h03),2) as basico from platemphist a,plaremunbase b "
                  F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                  F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
            End If
       Case Is = "24" 'BONIF. TURNO NOCHE
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h20),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "25" 'H.E. HASTA DECIMA HORA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h21),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "26" 'H.E. HASTA ONCEAVA HORA
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h22),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "27" 'EXTRAS 3PRA L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h23),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "28" 'EXT. NOCHE 2PR L-S
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h24),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "35" 'EXT. NOCHE 3PRA L-S (29)
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h25),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "30" 'SOBRETASA NOCHE(CONSTRUCCION CIVIL)
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vtiptrab & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsF01, Sql$)) Then
               mFactor = rsF01!factor
               If rsF01.State = 1 Then rsF01.Close
               F01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h07),2) as basico from platemphist a,plaremunbase b "
               F01 = F01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               F01 = F01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               F01 = F01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
End Select

End Function

Public Function V01(ByVal wcia As String, ByVal wcodpla As String, ByVal vtope As String, ByVal VCargo As String, ByVal Vano As Integer, ByVal Vmes As Integer, _
                    ByVal Mstatus As String, ByVal vTipoTra As String, BYVALconcepto As String) As String
'INGRESOS
Dim rsV01 As ADODB.Recordset
Dim mFactor As Currency
Dim fACTORfAMILIA As String

'FECHA DE MODIFICACION 08/01/2008
'If wcia = "05" Then fACTORfAMILIA = 2 Else fACTORfAMILIA = 1

If wcia = "05" Then fACTORfAMILIA = 1 Else fACTORfAMILIA = 1

mFactor = 0
V01 = ""
Select Case concepto
       Case Is = "02" 'ASIGNACION FAMILIAR
            V01 = "select round((b.importe/factor_horas)*(a.h12*" & fACTORfAMILIA & "),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "04" 'BONIFICACION T. SERVICIO
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "05" 'INCREMENTO AFP 10.23%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "06" 'INCREMENTO AFP 3%
            V01 = "select round((b.importe/CASE WHEN TIPO=02 THEN factor_horas+8 ELSE FACTOR_HORAS END)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "07" 'BONIFICACION COSTO DE VIDA
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
       Case Is = "14" 'VACACIONES
            V01 = "select round((b.importe/factor_horas)*(a.h12),2) as basico from platemphist a,plaremunbase b "
            V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
            V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
            V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
       Case Is = "20" 'BUC
            Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
            If (fAbrRst(rsV01, Sql$)) Then
               mFactor = rsV01!factor
               If rsV01.State = 1 Then rsV01.Close
               V01 = "select round(((b.importe/factor_horas)* " & mFactor & ") *(a.h12),2) as basico from platemphist a,plaremunbase b "
               V01 = V01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
               V01 = V01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
               V01 = V01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
            End If
       Case Is = "29"
       
        'MODIFICADO POR RICARDO HINOSTROZA
        'AGREGAR CALCULO DE HORARIO VACACIONAL
        'FEC MODIFICACION 10/04/2008
        
             V01 = "SELECT ROUND(SUM(I29)/6,2)  AS BASICO FROM PLAHISTORICO WHERE CIA ='" & wcia & "' AND STATUS<>'*' AND PLACOD='" & wcodpla & "' AND YEAR(FECHAPROCESO) =" & Vano & " AND (MONTH(FECHAPROCESO) BETWEEN 1 AND 3)"
       
End Select
'Debug.Print concepto
'Debug.Print SQL$
End Function

Public Function G01(ByVal wcia As String, ByVal wcodpla As String, ByVal vtope As String, ByVal VCargo As String, _
                    ByVal vTipoTra As String, ByVal Vano As Integer, ByVal Vmes As Integer, ByVal Mstatus As String, concepto As String) As String      'INGRESOS
Dim rsG01 As ADODB.Recordset
Dim mFactor As Currency
Dim mh As Integer
mFactor = 0
If Val(Mid(VFProceso, 4, 2)) = 7 Then mh = 7 Else mh = 5
G01 = ""
If vTipoTra <> "05" Then
    Select Case concepto
           Case Is = "02" 'ASIGNACION FAMILIAR
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "04" 'BONIFICACION T. SERVICIO
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "05" 'INCREMENTO AFP 10.23%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "06" 'INCREMENTO AFP 3%
                G01 = "select round((((b.importe/CASE WHEN TIPO='02' THEN factor_horas+8 ELSE FACTOR_HORAS END)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "07" 'BONIFICACION COSTO DE VIDA
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='" & concepto & "' and b.status<>'*'"
           Case Is = "15" 'GRATIFICACION
                G01 = "select round((((b.importe/factor_horas)*240)/6)*(a.h21),2) as basico from platemphist a,plaremunbase b "
                G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
    End Select
Else
   Select Case concepto
          Case Is = "15" 'GRATIFICACION
               Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='" & concepto & "' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
               If (fAbrRst(rsG01, Sql$)) Then
                  mFactor = rsG01!factor
                  If rsG01.State = 1 Then rsG01.Close
                  G01 = "select round(((((b.importe/factor_horas)*8)* " & mFactor & ") / " & mh & ") * a.h14 ,2) as basico from platemphist a,plaremunbase b "
                  G01 = G01 & "where a.cia='" & wcia & "' and a.status='" & Mstatus & "' and a.placod='" & wcodpla & "' "
                  G01 = G01 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " "
                  G01 = G01 & "and b.cia=a.cia and b.placod=a.placod and b.concepto='01' and b.status<>'*'"
               End If
   End Select
End If
End Function

Public Function F02(ByVal wcia As String, ByVal wcodpla As String, ByVal wcodafp As String, ByVal vtope As String, ByVal VCargo As String, _
                    ByVal sn_essaludvida As String, ByVal sn_sindicato As String, ByVal VFechaJub As String, ByVal wtipodoc As Boolean, _
                    ByVal VTipobol As String, ByVal Mstatus As String, ByVal Vano As Integer, ByVal Vmes As Integer, ByVal VHoras As Integer, _
                    ByVal VSemana As String, ByVal sn_quinta As Boolean, concepto As String) As String   'DEDUCCIONES
Dim rsF02 As ADODB.Recordset
Dim rsF02afp As ADODB.Recordset
Dim F02str As String
Dim rsTope As ADODB.Recordset
Dim mFactor As Currency
Dim mperiodoafp As String
Dim vNombField As String
Dim mtope As Currency
Dim mincremento As Currency
Dim difmincremento As Currency
Dim mproy As Currency
Dim MUIT As Currency
Dim mgra As Integer
Dim msemano As Integer
Dim mpertope As Integer
Dim J As Integer
Dim conceptosremu As String
Dim snmanual As Byte, snfijo As Byte
Dim cptoincrementos As String
Dim IMPORTEESVMES As Currency

mFactor = 0
F02 = ""
mtope = 0

If (concepto <> "04" Or Trim(wcodafp) = "01" Or Trim(wcodafp) = "" Or Trim(wcodafp) = "02") And concepto <> "11" And concepto <> "13" Then  'SIN AFP
   If Not IsDate(VFechaJub) Then
    Sql$ = "select deduccion,adicional,status from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and deduccion<>0 and status<>'*'"
  
    If (fAbrRst(rsF02, Sql$)) Then
       If Not IsNull(rsF02!deduccion) Then
          If rsF02!deduccion <> 0 Then mFactor = rsF02!deduccion: snmanual = IIf(rsF02!adicional = "S", 0, 1): snfijo = IIf(rsF02!status = "F", 1, 0)
       End If
    End If
    
    If rsF02.State = 1 Then rsF02.Close
    
    If snfijo = 1 And snmanual = 0 And ((sn_essaludvida = 1 And concepto = "06") Or (sn_sindicato = 1 And concepto = "08")) Then
        Call Acumula_Mes(wcia, wcodpla, Vano, Vmes, concepto, "D")
        If wtipodoc = True Then
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
                F02 = " SELECT " & mFactor
            End If
        Else
            If macui = mFactor Then
                F02 = " SELECT " & 0
            Else
                F02 = " SELECT " & mFactor / 2
            End If
        End If
        Exit Function
    End If
    
    If mFactor <> 0 Then
       Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
       If (fAbrRst(rsF02, Sql$)) Then
          rsF02.MoveFirst
          F02str = ""
          Do While Not rsF02.EOF
             F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
             rsF02.MoveNext
          Loop
          F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
          If rsF02.State = 1 Then rsF02.Close
          Call Acumula_Mes(wcia, wcodpla, Vano, Vmes, concepto, "D")
          F02 = "select round(((" & F02str & " + " & macui & ") * " & mFactor & " /100)-" & macus & ",2) as deduccion from platemphist "
          F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' "
          F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
       End If
    End If
   End If
ElseIf Trim(concepto) = "11" And Trim(wcodafp) <> "" Then 'AFP
   If Not IsDate(VFechaJub) Then
   If Trim(wcodafp) = "01" Or Trim(wcodafp) = "02" Then GoTo AFP
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"

    If (fAbrRst(rsF02, Sql$)) Then
       rsF02.MoveFirst
       F02str = ""
       Do While Not rsF02.EOF
          F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
          
          rsF02.MoveNext
       Loop
       
       F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
       
       If rsF02.State = 1 Then rsF02.Close
       mperiodoafp = Format(Vano, "0000") & Format(Vmes, "00")
       Sql$ = "select afp01,afp02,afp03,afp04,afp05,tope from  plaafp where periodo='" & mperiodoafp & "' and codafp='" & wcodafp & "' and status<>'*' and cia='" & wcia & "'"
    
       If Not (fAbrRst(rsF02, Sql$)) Then
          MsgBox "No se Encuentran Factores de Calculo para AFP", vbCritical, "Calculo de Boleta"
          Exit Function
       End If
       Sql$ = Acumula_Mes_Afp(wcia, wcodpla, Vano, Vmes, concepto, "D")
       If (fAbrRst(rsF02afp, Sql$)) Then
          For J = 1 To 5
              vNombField = " as D11" & Format(J, "0")
              mFactor = rsF02(J - 1)
              If J = 2 Then
                 If manos > 64 Then mFactor = 0
                 Call Acumula_Mes_Afp112(wcia, wcodpla, Vano, Vmes, concepto, "D")
                 mtope = macui
                 Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' " _
                      & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
                      
                 If (fAbrRst(rsTope, Sql$)) Then mtope = mtope + rsTope!tope
                 
                 If wtipodoc = False Then mtope = mtope + rsTope!tope
                 
                 If mtope > rsF02!tope Then mtope = rsF02!tope
                                 
                 If rsTope.State = 1 Then rsTope.Close
                 
                 If wtipodoc = False Then
                    F02 = F02 & "round((((" & mtope & ") * " & mFactor & " /100)-" & macus & ")/2,2) "
                 Else
                    F02 = F02 & "round(((" & mtope & ") * " & mFactor & " /100)-" & macus & ",2) "
                 End If
                 F02 = F02 & vNombField & ","
              Else
                 If Not IsNull(rsF02afp(0)) Then
                    If wtipodoc Then
                    F02 = F02 & "round(((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & ",2) "
                    F02 = F02 & vNombField & ","
                    Else
                        F02 = F02 & "round((((" & F02str & " + " & rsF02afp(0) & ") * " & mFactor & " /100)-" & rsF02afp(J) & "),2) "
                        F02 = F02 & vNombField & ","
                    End If
                 Else
                    If wtipodoc Then
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    Else
                        F02 = F02 & "round(((" & F02str & " ) * " & mFactor & " /100),2) "
                        F02 = F02 & vNombField & ","
                    End If
                 End If
              End If
          Next J
          If rsF02afp.State = 1 Then rsF02afp.Close
          If rsF02.State = 1 Then rsF02.Close
       End If
       F02 = Mid(F02, 1, Len(Trim(F02)) - 1)
       F02 = "select " & F02
       F02 = F02 & " from platemphist "
       F02 = F02 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' "
       F02 = F02 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
AFP:
    End If
   End If
               'de 13 lo cambie a 131 para que no entre
ElseIf concepto = "13" And VTipobol <> "03" Then 'Quinta Categoria

    'PREGUNTAMOS SI ESTA AFECTO O NO
    If VTipobol = "04" Then GoTo quinta
    If VTipobol = "05" Then GoTo quinta
    If Not sn_quinta Then GoTo quinta
    
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='D' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF02, Sql$)) Then
      rsF02.MoveFirst
      F02str = ""
      conceptosremu = ""
      cptoincrementos = ""
      Do While Not rsF02.EOF
         F02str = F02str & "i" & Trim(rsF02!cod_remu) & "+"
'         If Trim(rsF02!cod_remu) <> "01" And Trim(rsF02!cod_remu) <> "14" And Trim(rsF02!cod_remu) <> "12" Then
'            conceptosremu = conceptosremu & "'" & Trim(rsF02!cod_remu) & "',"
'            cptoincrementos = cptoincrementos & "I" & Trim(rsF02!cod_remu) & "+"
'         End If
         rsF02.MoveNext
      Loop

'      conceptosremu = Mid(conceptosremu, 1, Len(Trim(conceptosremu)) - 1)
'      cptoincrementos = Mid(cptoincrementos, 1, Len(Trim(cptoincrementos)) - 1)
      
      F02str = "(" & Mid(F02str, 1, Len(Trim(F02str)) - 1) & ")"
      If rsF02.State = 1 Then rsF02.Close
      MUIT = 0
      mtope = 0
      mpertope = 0
      mgra = 0
      msemano = 0
      mincremento = 0
      If vTipoTra <> "01" Then
         Sql$ = "select isnull(max(semana),0) as ultima from plasemanas where cia='" & wcia & "' and ano='" & Format(Vano, "0000") & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mpertope = rsF02(0): msemano = rsF02(0)
      Else
         If Vmes > 6 Then mpertope = 12 Else mpertope = 13
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      'ACUMULADO DE TODOS INGRESOS
      Call Acumula_Ano(wcia, wcodpla, Vano, Vmes, concepto, "D")
      mtope = macui
'      cptoincrementos = F02str
      
      'OBTENER EL INCREMENTO
      If Len(Trim(cptoincrementos)) > 0 Then
        Sql$ = "select " & cptoincrementos & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' " _
             & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
        If (fAbrRst(rsF02, Sql$)) Then mincremento = rsF02!tope
        If rsF02.State = 1 Then rsF02.Close
      End If
      
      Sql$ = "select " & F02str & " as tope from platemphist where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' " _
           & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      If (fAbrRst(rsF02, Sql$)) Then mtope = mtope + rsF02!tope
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select concepto,moneda,sum((importe/factor_horas)) as base from plaremunbase a,plaafectos b where a.cia='" & wcia & "' " _
           & "and placod='" & wcodpla & "' and a.status<>'*' and b.tipo='D' AND B.CIA='" & wcia & "' and b.codigo='" & concepto & "' and b.tboleta='" & VTipobol & "' and b.status<>'*' and a.concepto=b.cod_remu " _
           & "Group By Placod,a.concepto,a.moneda"
      
      If (fAbrRst(rsF02, Sql$)) Then
        Do While Not rsF02.EOF
            mproy = mproy + rsF02!base
            rsF02.MoveNext
        Loop
        'mproy = Round(mproy, 2)
      End If
      If rsF02.State = 1 Then rsF02.Close
      
      Sql$ = "select uit from plauit where ano='" & Format(Vano, "0000") & "' and moneda='S/.' and status<>'*'"
      If (fAbrRst(rsF02, Sql$)) Then MUIT = rsF02!uit
      If rsF02.State = 1 Then rsF02.Close
      
      If vTipoTra = "05" Then
         Sql$ = "select factor from platasaanexo where modulo='02' and tipomovimiento='" & Trim(vtope) & "' and codinterno='20' and tipotrab='" & vTipoTra & "' and cargo='" & Trim(VCargo) & "' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mFactor = rsF02!factor
         If rsF02.State = 1 Then rsF02.Close
      
         Sql$ = "select concepto,moneda,importe/factor_horas as base from plaremunbase where cia='" & wcia & "' and placod='" & wcodpla & "' and concepto='01' and status<>'*'"
         If (fAbrRst(rsF02, Sql$)) Then mproy = mproy + Round(rsF02!base * mFactor, 2)
         If rsF02.State = 1 Then rsF02.Close
      End If
      If vTipoTra = "01" Then
         If wtipodoc = True Then
            If Vmes < 12 Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Vmes + 1) Else mproy = 0
         Else
            If Vmes = 12 Then
               mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes)) + Round((((mproy * VHoras) + difmincremento) / 2), 2)
'               --Vmes = Vmes - 1
               'mproy = 0
            Else
'                SQL = "select sum(importe) from plaremunbase b where b.cia='" & wcia & "' and b.placod='" & Trim(wcodpla) & "' and b.concepto in (" & conceptosremu & ") and b.status<>'*'"
'                difmincremento = 0
'                If (fAbrRst(rsF02, SQL$)) Then difmincremento = IIf(IsNull(rsF02(0)), 0, rsF02(0))
'                If rsF02.State = 1 Then rsF02.Close
                
                '{>MA<} 090407
              mproy = (((mproy * VHoras) + difmincremento) * (mpertope - Vmes + 1)) + Round((((mproy * VHoras) + difmincremento) / 2), 2)
              'mproy = ((mproy * VHoras) + mincremento) * (mpertope - Vmes + 1)
            End If
         End If
      Else
         mgra = Busca_Grati(wcia, wcodpla, Vano, Vmes)
         If vTipoTra = "05" Then
            Sql$ = "select importe/factor_horas as base,b.factor  from plaremunbase a,platasaanexo b where a.cia='" & wcia & "' and a.placod='" & wcodpla & "'  and a.concepto='01' " _
                 & "and a.status<>'*' and b.cia='" & wcia & "' and b.tipomovimiento='01' and b.codinterno='15' and b.status<>'*' and b.tipotrab='" & vTipoTra & "' and b.cargo='" & Trim(VCargo) & "'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + ((rsF02!base * 8) * (rsF02!factor * mgra)) + ((rsF02!base * 8) * (mpertope - Val(VSemana)))
            If rsF02.State = 1 Then rsF02.Close
         Else
            Sql$ = "select importe/factor_horas as base from plaremunbase where Cia='" & wcia & "' and placod='" & wcodpla & "'  and concepto='01' and status<>'*'"
            If (fAbrRst(rsF02, Sql$)) Then mproy = ((mproy * VHoras) + mincremento) * (mpertope - Val(VSemana)) + (((mproy * 240) + mincremento) * mgra) + (rsF02!base * 8) * (mpertope - Val(VSemana))
            If rsF02.State = 1 Then rsF02.Close
            'mproy = (mproy * VHoras) * (mpertope - Val(VSemana)) + ((mproy * 240) * mgra) + (mproy * 8) * (mpertope - Val(VSemana))
         End If
      End If
      mtope = mtope + mproy
      If mtope > Round(MUIT * 7, 2) Then
         mtope = mtope - Round(MUIT * 7, 2)
         Select Case mtope
                Case Is < (Round(MUIT * 27, 2) + 1)
                     mFactor = Round(mtope * 0.15, 2)
                Case Is < (Round(MUIT * 54, 2) + 1)
                     mFactor = Round(((mtope - (MUIT * 27)) * 0.21) + (MUIT * 27) * 0.15, 2)
                Case Else
                     mFactor = Round(((mtope - (MUIT * 54)) * 0.27) + ((MUIT * 54) - (MUIT * 27)) * 0.21, 2) + ((MUIT * 27) * 0.15)
         End Select
         
         If vTipoTra = "01" Then
            If wtipodoc = True Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round(((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1)), 2)"
            End If
         Else
            If VTipobol = "02" Then
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (12 - " & Vmes & " + 1), 2)"
            Else
               F02 = "Select Round((" & mFactor & " - " & macus & ") / (" & msemano & " - " & Val(VSemana) & " + 1), 2)"
            End If
         End If
      End If
   End If
quinta:
   
ElseIf concepto = "15" Then
       MsgBox "666"
End If
macui = 0: macus = 0
End Function

Private Sub Acumula_Mes(ByVal wcia As String, ByVal wcodpla As String, ByVal Vano As Integer, ByVal Vmes As Integer, concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If concepto = "06" And tipo = "D" Then
        Sql$ = "select 'D06' AS COD_REMU"
    Else
        Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    End If
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
        If concepto = "06" And tipo = "D" Then
            mcad = mcad & Trim(RsAcumula!cod_remu) & "+"
        Else
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
        End If
          RsAcumula.MoveNext
       Loop
       
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & wcodpla & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Acumula_Mes_Afp(ByVal wcia As String, ByVal wcodpla As String, ByVal Vano As Integer, ByVal Vmes As Integer, concepto As String, tipo As String) As String
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset

macui = 0: macus = 0
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='01'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d111) as ded1, " _
              & "sum(d112) as ded2, " _
              & "sum(d113) as ded3, " _
              & "sum(d114) as ded4, " _
              & "sum(d115) as ded5 "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & wcodpla & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and "
       SqlAcu = SqlAcu & "proceso in('01','02','03','04','05')"
    End If
If RsAcumula.State = 1 Then RsAcumula.Close
Acumula_Mes_Afp = SqlAcu
End Function

Private Sub Acumula_Mes_Afp112(ByVal wcia As String, ByVal wcodpla As String, ByVal Vano As Integer, ByVal Vmes As Integer, concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 3
    If VTipobol = "02" And I <> 2 Then I = I + 1
    If VTipobol <> "02" And I = 2 Then I = I + 1
    If I > 3 Then Exit For
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop
       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(d112) as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & wcodpla & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " and proceso='" & mtb & "'"
       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub

Private Sub Acumula_Ano(ByVal wcia As String, ByVal wcodpla As String, ByVal Vano As Integer, ByVal Vmes As Integer, concepto As String, tipo As String)
Dim mcad As String
Dim SqlAcu As String
Dim RsAcumula As New Recordset
Dim rs2 As ADODB.Recordset
Dim mtb As String
macui = 0: macus = 0
For I = 1 To 4
    If I = 1 Then mtb = "01"
    If I = 2 Then mtb = "02"
    If I = 3 Then mtb = "03"
    If I = 4 Then mtb = "04"
    Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='" & tipo & "' and tboleta='" & mtb & "'  and  codigo='" & concepto & "' and status<>'*'"
    If (fAbrRst(RsAcumula, Sql$)) Then
       RsAcumula.MoveFirst
       mcad = ""
       Do While Not RsAcumula.EOF
          mcad = mcad & "i" & Trim(RsAcumula!cod_remu) & "+"
          RsAcumula.MoveNext
       Loop

       mcad = "(" & Mid(mcad, 1, Len(Trim(mcad)) - 1) & ")"
       SqlAcu = "select sum(" & mcad & ") as ing,sum(" & Trim(tipo) & Trim(concepto) & ") as ded "
       SqlAcu = SqlAcu & "from plahistorico "
       SqlAcu = SqlAcu & "where cia='" & wcia & "' and status<>'*' and placod='" & wcodpla & "' "
       SqlAcu = SqlAcu & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)<=" & Vmes & " and proceso='" & mtb & "'"

       If (fAbrRst(rs2, SqlAcu)) Then
          If Not IsNull(rs2!ing) Then macui = macui + rs2!ing
          If Not IsNull(rs2!ded) Then macus = macus + rs2!ded
       End If
    End If
Next
If RsAcumula.State = 1 Then RsAcumula.Close
End Sub
Private Function Busca_Grati(ByVal wcia As String, ByVal wcodpla As String, ByVal Vano As Integer, ByVal Vmes As Integer) As Integer
Dim rsGrati As New Recordset
Select Case Vmes
       Case Is = 1, 2, 3, 4, 5, 6
            Busca_Grati = 2
       Case Is = 7
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & wcodpla & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 1 Else Busca_Grati = 2
                 
       Case Is = 8, 9, 10, 11
            Busca_Grati = 1
       Case Is = 12
            Sql$ = "select * from plahistorico where cia='" & wcia & "' and placod='" & wcodpla & "' and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & " " _
                 & "and proceso='03' and status<>'*'"
            If (fAbrRst(rsGrati, Sql$)) Then Busca_Grati = 0 Else Busca_Grati = 1
End Select
End Function

Public Function F03(ByVal wcia As String, ByVal wcodpla As String, ByVal VArea As String, ByVal VTipobol As String, _
                    ByVal Vano As Integer, ByVal Vmes As Integer, ByVal Mstatus As String, concepto As String) As String 'APORTACIONES
Dim rsF03 As ADODB.Recordset
Dim rscalculo As ADODB.Recordset
Dim F03str As String
Dim mFactor As Currency
mFactor = 0
F03 = ""
If concepto = "03" Then
   Sql$ = "select senati from cia where cod_cia='" & wcia & "' and status<>'*'"
   If Not (fAbrRst(rsF03, Sql$)) Then Exit Function
   If Trim(rsF03!senati) <> "S" Then Exit Function
   If rsF03.State = 1 Then rsF03.Close
   wciamae = Determina_Maestro("01044")
   Sql$ = "Select * from maestros_2 where cod_maestro2='" & Trim(VArea) & "' and status<>'*'"
   Sql$ = Sql$ & wciamae
   If (fAbrRst(rsF03, Sql$)) Then
      If rsF03!flag7 <> "S" Then Exit Function
   Else
      Exit Function
   End If
   If rsF03.State = 1 Then rsF03.Close
End If

Sql$ = "select aportacion from placonstante where cia='" & wcia & "' and tipomovimiento='03' and codinterno='" & concepto & "' and aportacion<>0 and status<>'*'"
If (fAbrRst(rsF03, Sql$)) Then
   If Not IsNull(rsF03!aportacion) Then
      If rsF03!aportacion <> 0 Then mFactor = rsF03!aportacion
   End If
End If
If rsF03.State = 1 Then rsF03.Close
If mFactor <> 0 Then
   Sql$ = "select cod_remu from plaafectos where cia='" & wcia & "' and tipo='A' and tboleta='" & VTipobol & "'  and  codigo='" & concepto & "' and status<>'*'"
   If (fAbrRst(rsF03, Sql$)) Then
      rsF03.MoveFirst
      F03str = ""
      Do While Not rsF03.EOF
         F03str = F03str & "i" & Trim(rsF03!cod_remu) & "+"
         rsF03.MoveNext
      Loop
      
      F03str = "(" & Mid(F03str, 1, Len(Trim(F03str)) - 1) & ")"
      If rsF03.State = 1 Then rsF03.Close
      
      Call Acumula_Mes(wcia, wcodpla, Vano, Vmes, concepto, "A")
      
      F03 = "select (" & F03str & " + " & IIf(Val(macui) = 0, 0, macui) & ") , (" & mFactor & " /100)," & IIf(Val(macus) = 0, 0, macus) & " from platemphist "
      F03 = F03 & "where cia='" & wcia & "' and status='" & Mstatus & "' and placod='" & wcodpla & "' "
      F03 = F03 & "and year(fechaproceso)=" & Vano & " and month(fechaproceso)=" & Vmes & ""
      
      If (fAbrRst(rscalculo, F03)) Then
        If rscalculo(0) > sueldominimo Then
            F03 = " SELECT " & Round((rscalculo(0) * (mFactor / 100)) - macus, 2)
        Else
            F03 = "SELECT " & Round((sueldominimo * (mFactor / 100)) - macus, 2)
        End If
      End If
      
   End If
End If

macui = 0: macus = 0
End Function
Public Function Descarga_ctaCte(ByVal wcia As String, ByVal wcodpla As String, ByVal wtipodoc As Boolean, ByVal wcodaux As String, tipobol As String, fecha As String, sem As String, tiptrab As String, importe As Currency) As Boolean
Dim Sql As String
Dim RX As ADODB.Recordset
'SQL = "select * from plactacte where cia='" & wcia & "' and placod='" & Trim(wcodpla) & "' and importe-pago_acuenta>0 and status<>'*' order by fecha"
'If (fAbrRst(rs, SQL$)) Then rs.MoveFirst
'Do While Not rs.EOF
'   saldo = (rs!importe - rs!pago_acuenta)
'   If saldo >= importe Then
'      If saldo = importe Then
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      Else
'         'SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & wcia & "' and numinterno='" & RS!numinterno & "' and tipo='" & RS!tipo & "' and status<>'*'"
'         SQL = "update plactacte set pago_acuenta=pago_acuenta+" & importe & " where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'         cn.Execute SQL
'      End If
      QUINCENA = 0
      For I = 0 To MAXROW - 1
        If ArrDsctoCTACTE(0, I) = Trim(wcodpla) Then
            If ArrDsctoCTACTE(2, I) > 0 Then
            
                Sql = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
                Sql = Sql & "insert into plabolcte values('" & wcia & "','" & UCase(Trim(wcodpla)) & "','" & tipobol & "','', " _
                & "'" & fecha & "','" & sem & "','" & tiptrab & "','" & wcodaux & "','" & ArrDsctoCTACTE(3, I) & "','" & fecha & "', " _
                & "'" & wmoncont & "'," & ArrDsctoCTACTE(2, I) & ",'','" & wuser & "'," & FechaSys & "," & ArrDsctoCTACTE(1, I) & "," & IIf(wtipodoc = True, 0, 1) & ")"
                
                cn.Execute Sql
            End If
        
            Sql = "UPDATE plactacte set pago_acuenta=pago_acuenta + " & CStr(ArrDsctoCTACTE(2, I)) + " WHERE cia='" & wcia & "' and placod='" & Trim(wcodpla) & "' and id_doc='" & CStr(ArrDsctoCTACTE(1, I)) & "' and status<>'*'"
            cn.Execute Sql
            
            Sql = "UPDATE plactacte set fecha_cancela='" & fecha & "' WHERE cia='" & wcia & "' and placod='" & Trim(wcodpla) & "' and id_doc='" & ArrDsctoCTACTE(1, I) & "' and status<>'*' AND IMPORTE=PAGO_ACUENTA"
            cn.Execute Sql
                
        End If
      Next
      'Exit Do
'   Else
'      SQL = "update plactacte set pago_acuenta=pago_acuenta+" & saldo & ",fecha_cancela='" & Format(fecha, FormatFecha) & "' where cia='" & wcia & "' and CODAUXINTERNO='" & rs!codauxinterno & "' and tipo='" & rs!tipo & "' and status<>'*'"
'      cn.Execute SQL
'
'      SQL = "SET DATEFORMAT " & Coneccion.FormatFechaSql & " "
'      SQL = SQL & "insert into plabolcte values('" & wcia & "','" & wcodpla & "','" & tipobol & "','', " _
'          & "'" & Format(fecha, FormatFecha) & "','" & sem & "','" & tiptrab & "','" & rs!codauxinterno & "','" & rs!tipo & "','" & Format(rs!fecha, FormatFecha) & "', " _
'          & "'" & wmoncont & "'," & saldo & ",'','" & wuser & "'," & FechaSys & ")"
'      cn.Execute SQL
'      importe = importe - saldo
'   End If
'   rs.MoveNext
'Loop
'rs.Close
End Function


Public Function MUESTRA_CUENTACORRIENTE(ByVal wcia As String, ByVal wcodpla As String, ByVal wtipodoc As Boolean) As Boolean
Dim RX As New ADODB.Recordset
Dim I As Integer
Dim sumporc As Currency
On Error GoTo CORRIGE
MAXROW = 0
'VSemana
'CON = "SELECT DES.MONTO FROM PLACTACTE CTA,PLADESCTA DES WHERE " & _
'"CTA.PLACOD='" & Trim(wcodpla) & "' AND (CTA.IMPORTE-CTA.PAGO_ACUENTA)>0 AND " & _
'"DES.CODAUXINTERNO=CTA.CODAUXINTERNO AND RIGHT(DES.FECHA,2)='" & _
'VSemana & "' AND CTA.STATUS<>'*'"
MUESTRA_CUENTACORRIENTE = False
If BoletaCargada Then Exit Function

If VTipobol = "01" Then
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(wcodpla) & "' AND a.STATUS<>'*'"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
Else
    con = "SELECT a.id_doc,a.tipo,a.importe,COALESCE(SUM(b.importe),0) as pago_acuenta,a.partes,"
    con = con & " a.partes-(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=0"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as saldomes,"
    con = con & "(SELECT COALESCE(SUM(importe),0) FROM plabolcte WHERE cia=a.cia and placod=a.placod and id_doc=a.id_doc and status!='*' AND sn_quincena=1"
    con = con & " AND MONTH(fechaproceso)=" & Month(CDate(VFProceso)) & " AND YEAR(fechaproceso)=" & Year(CDate(VFProceso)) & ") as quincena"
    con = con & " FROM PLACTACTE a LEFT OUTER JOIN PLABOLCTE b ON"
    con = con & " (b.cia=a.cia and b.placod=a.placod and b.id_doc=a.id_doc and b.status!='*') WHERE a.cia='" & wcia & "' and a.placod='" & Trim(wcodpla) & "' AND a.STATUS<>'*' and a.sn_grati=1"
    con = con & " GROUP BY a.cia,a.placod,a.id_doc,a.tipo,a.importe,a.partes"
End If

RX.Open con, cn, adOpenStatic, adLockReadOnly

   If RX.RecordCount > 0 Then

      rsdesadic.MoveFirst
      Erase ArrDsctoCTACTE
      MAXROW = 0
      Do While Not rsdesadic.EOF
         If rsdesadic("CODIGO") = "07" Then
            Do While Not RX.EOF
                ReDim Preserve ArrDsctoCTACTE(0 To 4, 0 To MAXROW)
                
                ArrDsctoCTACTE(0, MAXROW) = wcodpla
                ArrDsctoCTACTE(1, MAXROW) = RX!id_doc
                ArrDsctoCTACTE(3, MAXROW) = RX!tipo
                ArrDsctoCTACTE(4, MAXROW) = 0
                
                If RX("partes") >= (RX("IMPORTE") - RX("PAGO_ACUENTA")) Then
                   rsdesadic("MONTO") = rsdesadic("MONTO") + (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                   If wtipodoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("IMPORTE") - RX("PAGO_ACUENTA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = (RX("IMPORTE") - RX("PAGO_ACUENTA"))
                Else
                   rsdesadic("MONTO") = rsdesadic("MONTO") + IIf(RX("partes") = 0, 0, RX("partes") - RX("QUINCENA"))
                   If wtipodoc = False Then ArrDsctoCTACTE(2, MAXROW) = Round((RX("partes") - RX("QUINCENA")) / 2, 2) Else ArrDsctoCTACTE(2, MAXROW) = RX("partes") - RX("QUINCENA")
                End If
                
'                If wTipoDoc = False Then
'                    If rsdesadic("MONTO") > 0 Then
'                        rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
'                    End If
'                End If
                
                MAXROW = MAXROW + 1
                
                RX.MoveNext
            Loop
            
             If wtipodoc = False Then
                If rsdesadic("MONTO") > 0 Then
                    rsdesadic("MONTO") = Round(rsdesadic("MONTO") / 2, 2)
                End If
            End If
                
            sumporc = 0
            If rsdesadic("MONTO") > 0 Then
                For I = 0 To MAXROW - 1
                
                    ArrDsctoCTACTE(4, I) = Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    sumporc = sumporc + Round((ArrDsctoCTACTE(2, I) * 100) / rsdesadic("MONTO"), 2)
                    If I = MAXROW - 1 Then
                        If sumporc > 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) - (sumporc - 100)
                        ElseIf sumporc < 100 Then
                            ArrDsctoCTACTE(4, I) = ArrDsctoCTACTE(4, I) + (100 - sumporc)
                        End If
                    End If
                Next
            End If
         End If
         rsdesadic.MoveNext
      Loop
   End If

RX.Close

MUESTRA_CUENTACORRIENTE = True
Exit Function
CORRIGE:
'MsgBox "Error :" & Err.Description, vbCritical, Caption

End Function


