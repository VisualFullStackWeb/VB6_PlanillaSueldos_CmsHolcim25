Attribute VB_Name = "Formulas"
Global Const iNUEVO = 0
Global Const iGRABAR = 1
Global Const iCANCELAR = 2
Global Const iSALIR = 3
Global Const iMAS = 4
Global Const iMENOS = 5
Global Const iOK = 6
Global Const iBAD = 7
Global Const iCHCKINAC = 8
Global Const iCHCKACT = 9
Global Const iCHCK = 10
Global Const iIMPRIMIR = 11
Global Const iGRAFICO = 12
'*********************************

Global Const HORAS_EMPLEADO = 240
Global Const HORAS_OBRERO = 48
Global Const hORAS_X_DIA = 8
Global Const DIAS_TRABAJO = 30


Public Function Apostrofe(ByVal strSql As Variant) As String
    Dim i As Long
    Dim strConcatenar As String
    Dim strResultado As String
    
    For i = 1 To Len(strSql)
        If Mid(strSql, i, 1) = "'" Then
            strConcatenar = "''"
        Else
            strConcatenar = Mid(strSql, i, 1)
        End If
        
        strResultado = strResultado + strConcatenar
    Next i
    Apostrofe = strResultado
    
End Function
