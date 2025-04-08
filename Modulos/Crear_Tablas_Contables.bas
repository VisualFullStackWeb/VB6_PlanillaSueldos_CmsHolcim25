Attribute VB_Name = "Crear_Tablas_Contables"
Option Explicit
Dim rs_Mcontable As ADODB.Recordset
Dim rs_Mcontable2 As ADODB.Recordset
Sub Proceso_Integral()
    If Verifica_Existencia_Tablas = False Then
        Call Proceso_Inicio_Tablas
    End If
End Sub
Public Function getCC_UltimaPlanilla(cod_cia As String, codEmp As String, ayo As Integer, mes As Integer) As String

 getCC_UltimaPlanilla = ""

On Error GoTo laCagada

Dim resultado As String

Dim query As String

Dim rs_tmp As New ADODB.Recordset

query = "select ccosto1 from plahistorico where cia ='" & cod_cia & "' and placod = '" & codEmp & "' and  "

query = query & "YEAR(fechaproceso)=" & ayo & " and month(fechaproceso)=" & mes & " and status<>'*' and semana in (select MAX(semana) from dbo.plahistorico where cia ='" & cod_cia & "' and placod = '" & codEmp & "' and  YEAR(fechaproceso)=" & ayo & "  and month(fechaproceso)=" & mes & " and status<>'*')"

rs_tmp.Open query, cn, adOpenStatic, adLockOptimistic

If rs_tmp.RecordCount > 0 Then

    resultado = (rs_tmp(0))

End If

rs_tmp.Close

 

query = "select flag1 from maestros_2 where ciamaestro='" & cod_cia & "044'  and cod_maestro2='" & resultado & "'"

 

Set rs_tmp = New ADODB.Recordset

rs_tmp.Open query, cn, adOpenStatic, adLockOptimistic

If rs_tmp.RecordCount > 0 Then

    resultado = (rs_tmp(0))

End If

 

rs_tmp.Close

Set rs_tmp = Nothing

 

 getCC_UltimaPlanilla = RTrim(resultado)

Exit Function

laCagada:

MsgBox Err.Description

Exit Function

Resume

 

End Function


Sub Proceso_Inicio_Tablas()
    Dim i_Contador As Integer
    Call Crea_Tablas_Contable_Maestras
    Call Recupera_Tipos_Trabajadores
    Set rs_Mcontable2 = Crear_Plan_Contable.rs_PlanCont_Pub
    Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
    rs_Mcontable2.MoveFirst
    Do While Not rs_Mcontable2.EOF
        For i_Contador = 1 To 3
            Select Case i_Contador
                Case 1: Call Recupera_Conceptos_Ingresos(wcia, 1, "")
                Case 2: Call Recupera_Conceptos_Deduccion(wcia, 1, "")
                Case 3: Call Recupera_Conceptos_Aportes(wcia, 1, "")
            End Select
            Set rs_Mcontable = Crear_Plan_Contable.rs_PlanCont_Pub
            Set Crear_Plan_Contable.rs_PlanCont_Pub = Nothing
            rs_Mcontable.MoveFirst
            Do While Not rs_Mcontable.EOF
                Select Case i_Contador
                    Case 1:
                        Call Graba_Informacion_Contable_Maestros1("I" & rs_Mcontable!codinterno, _
                        rs_Mcontable!Descripcion, "I" & rs_Mcontable!codinterno & rs_Mcontable2!cod_maestro2)
                    Case 2
                        Call Graba_Informacion_Contable_Maestros1("D" & rs_Mcontable!codinterno, _
                        rs_Mcontable!Descripcion, "D" & rs_Mcontable!codinterno & rs_Mcontable2!cod_maestro2)
                    Case 3
                        Call Graba_Informacion_Contable_Maestros1("A" & rs_Mcontable!codinterno, _
                        rs_Mcontable!Descripcion, "A" & rs_Mcontable!codinterno & rs_Mcontable2!cod_maestro2)
                End Select
                rs_Mcontable.MoveNext
            Loop
            Set rs_Mcontable = Nothing
        Next i_Contador
        rs_Mcontable2.MoveNext
    Loop
    Set rs_Mcontable2 = Nothing
End Sub

