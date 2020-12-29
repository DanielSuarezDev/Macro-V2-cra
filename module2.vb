Public CadenaConexion As String
Public miBase As String
Public Conexion As ADODB.Connection
Public rsContab As ADODB.Recordset
Option Explicit
Sub abrirContabilizados()
 Dim Criterio, Sql As String
 Dim limpiarDatos, j As Long
 Dim ESPECIFICO  As String
 
    Call ConectarCont
    
    Set rsContab = New ADODB.Recordset
                Sheets("CONTABILIZADOS").Range("A12:AT1048576").ClearContents
                Criterio = Sheets("CONTABILIZADOS").Cells(4, 2)
                'ESPECIFICO = Sheets("CONTABILIZADOS").Cells(5, 2)
          'If Sheets("CONTABILIZADOS").Range("B5") = "TODOS" Then
              Sql = "SELECT * FROM [Contabilizados] WHERE [COD CONV] like '%" & Criterio & "%'"
          'ElseIf Sheets("CONTABILIZADOS").Range("B5") = "ED" Then
              'Sql = "SELECT * FROM [Contabilizados] WHERE [COD CONV] like  '%" & Criterio & "%'" And [NOMBRE CONVENIO] Like "'%" & ESPECIFICO & "%'"
          'End If
                            
                rsContab.Open Sql, Conexion
                Sheets("CONTABILIZADOS").Cells(9, 1).CopyFromRecordset rsContab
                limpiarDatos = Sheets("CONTABILIZADOS").Range("A" & Rows.Count).End(xlUp).Row
                For j = 9 To limpiarDatos
                    Sheets("CONTABILIZADOS").Cells(j, 1).Value = CLngLng(Sheets("CONTABILIZADOS").Cells(j, 1))
                Next j
    rsContab.Close
    Set rsContab = Nothing
    Conexion.Close
    Set Conexion = Nothing
    
End Sub
Sub ConectarCont()
miBase = "D:\PRUEBAS INFORMES\BASEtRABAJO.accdb"
CadenaConexion = "Provider=Microsoft.ACE.OLEDB.12.0; " & "data source=" & miBase & ";"
If Len(Dir(miBase)) = 0 Then
    MsgBox "La base que intenta conectar no se encuentra disponible", vbCritical
    Exit Sub
End If
Set Conexion = New ADODB.Connection
    If Conexion.State = 1 Then
        Conexion.Close
    End If
        Conexion.Open (CadenaConexion)
End Sub
Sub Insolvencias()
Dim LARGO, i As Long
Dim la, a As Long
LARGO = Sheets("ACTIVOS").Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To LARGO
If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("INSOLVENCIA").Range("C:C"), 1, 0)) Then
   If Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("INSOLVENCIA").Range("C:C"), 1, 0) Then
      Sheets("ACTIVOS").Cells(i, 52) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("INSOLVENCIA").Range("C:C"), 1, False)
      
      Else
      Sheets("ACTIVOS").Cells(i, 52) = "N/A"
   End If
End If
   
If Sheets("ACTIVOS").Cells(i, 52) > 1 Then
   Sheets("ACTIVOS").Cells(i, 46) = "INSOLVENCIA"
   MsgBox ("Por Favor Validar Insolvencia"), vbCritical, "Analisis Op"
End If

Sheets("ACTIVOS").Cells(i, 55) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A:AM"), 39, 0)
Sheets("ACTIVOS").Cells(i, 56) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("C:C"), 1, 0)




Next i


'///////////////CLIENTES ESPECIALES/////////////////////////////////////////
For i = 2 To LARGO
If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CLIENTES ESPECIALES").Range("E:E"), 1, 0)) Then
   If Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CLIENTES ESPECIALES").Range("E:E"), 1, 0) Then
      Sheets("ACTIVOS").Cells(i, 53) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CLIENTES ESPECIALES").Range("E:P"), 12, False)
      Else
      Sheets("ACTIVOS").Cells(i, 53) = "N/A"
   End If
End If
   
If Sheets("ACTIVOS").Cells(i, 53) > 1 Then
   Sheets("ACTIVOS").Cells(i, 46) = "ANALIZAR C.E"
   MsgBox ("Por Favor Validar Creditos Especiales"), vbCritical, "Analisis CSF"
End If

Next i


la = Sheets("CANCELADOS").Range("A" & Rows.Count).End(xlUp).Row
For a = 2 To la
Sheets("CANCELADOS").Cells(a, 50) = Application.VLookup(Sheets("CANCELADOS").Cells(a, 1), Sheets("ACTIVOS").Range("A:AT"), 46, 0)
Sheets("CANCELADOS").Cells(a, 51) = Application.VLookup(Sheets("CANCELADOS").Cells(a, 1), Sheets("ACTIVOS").Range("A:AK"), 37, 0)
Next a
'/////////////////////////////COLUMNAS FECHA FIN

'For i = 2 To LARGO
'If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A:AM"), 39, 0)) Then
'   If Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A:AM"), 39, 0) Then
 '     Sheets("ACTIVOS").Cells(i, 55) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("CANCELADOS").Range("A:AM"), 39, False)
      
 '     Else
  '    Sheets("ACTIVOS").Cells(i, 55) = "N/A"
      
  ' End If
'End If

'Next i

'///////////////////////////////CAN ANTERIORES/////////////
'For i = 2 To LARGO
'If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("C:C"), 1, 0)) Then
'   If Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("C:C"), 1, 0) Then
      
 '     Sheets("ACTIVOS").Cells(i, 56) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("VALIDA CANCELADOS").Range("C:C"), 1, False)
 '     Else
      
  '    Sheets("ACTIVOS").Cells(i, 56) = "N/A"
  ' End If
'End If

'Next i

'///////////////NOVEDADES/////////////////////////////////////////
For i = 2 To LARGO
If Not IsError(Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:A"), 1, 0)) Then
   If Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:A"), 1, 0) Then
      Sheets("ACTIVOS").Cells(i, 57) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:C"), 3, False)
      Sheets("ACTIVOS").Cells(i, 58) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:E"), 5, False)
      Sheets("ACTIVOS").Cells(i, 59) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:H"), 8, False)
      Sheets("ACTIVOS").Cells(i, 60) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:I"), 9, False)
      Sheets("ACTIVOS").Cells(i, 61) = Application.VLookup(Sheets("ACTIVOS").Cells(i, 1), Sheets("NOVEDADES").Range("A:AA"), 27, False)
      Else
      Sheets("ACTIVOS").Cells(i, 57) = "N/A"
      Sheets("ACTIVOS").Cells(i, 58) = "N/A"
      Sheets("ACTIVOS").Cells(i, 59) = "N/A"
      Sheets("ACTIVOS").Cells(i, 60) = "N/A"
      Sheets("ACTIVOS").Cells(i, 61) = "N/A"
   End If
End If
   


Next i




End Sub