Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports ZSQ_DOFIS.c_PA
Public Class f_PA
#Region "Parametros de conex y variables comunes"
    Dim o_conex As Object
    Dim p_server, p_basdat, p_usuari, p_passwr As String
    Dim Conexion As String
    'Variables de select,insert,delete 
    Dim p_ssql, p_dsql, p_isql As String
    ' cursores para recuperar el rec de ado
    Dim o_record_sql_sel, o_record_sql_ins, o_record_sql_del As Object
    Dim c, f, i As Integer
    Dim w_linea, w_valor, w_nfile As String
    Dim b_filtxt, c_filtxt As System.IO.StreamWriter
    Dim PA As String
    Dim w_ruta As String
    'declaracion de longitud del campo  
    Dim w_lng As Byte
    'Array de ubigeo
    Dim w_ubig As String()
    Dim w_ubic As String()
    'Declaramos la clase que tiene nuestras funciones  
    Dim cl_PA As New c_PA()
#End Region

    Private Sub btnListarPA_Click(sender As System.Object, e As System.EventArgs) Handles btnListarPA.Click
        PA = Convert.ToString(cboListaPA.Text)

        Select Case PA
            Case "PAIT0027"
                f_pait0027()
            Case "PAIT0037"
                f_pait0037()
            Case "PAIT0045"
                f_pait0045()
            Case "PAIT0045_1"
                f_pait0045_1()
            Case "PAIT0050"
                f_pait0050()
            Case "PAIT2006"
                f_pait2006()
            Case "PAIT2006_1"
                f_pait2006_1()
            Case "PAIT2012"
                f_pait2012()
             
        End Select

    End Sub

#Region "Generacion_textos_OM_PA_PI"
    '"--------------------------------------------------------
    '" Rutina de carga de distribución de C.Costo
    '"---------------------------------------------------------
    Sub f_pait0027()
        On Error GoTo ErrorHandler

        ' 1.-Retorna la conex a la base de datos
        o_conex = cl_PA.conexiondb()

        p_ssql =
            "SELECT " & _
            "MT.IDUSR  [PERNR], " & _
            "P1.FE_INGR_EMPR  [BEGDA], " & _
             "(SELECT FE_FINA_CONT from dbo.TDCONT_TRAB AS TC " & _
             "WHERE  (CO_EMPR = P1.CO_EMPR) AND (CO_TRAB = P1.CO_TRAB)  " & _
             "AND (ST_ULTI_CONT = 'S'))   [ENDDA], " & _
            "'01' [KSTAR], " & _
            "'300' [BURKS_01], " & _
            "C1.CO_CENT_COST  [KOSTL_01], " & _
            "C1.PC_DIST  [PROZT_01], " & _
            "'' [BURKS_02], " & _
            "'' [KOSTL_02],  " & _
            "'' [PROZT_02],  " & _
            "'' [BURKS_03],  " & _
            "'' [KOSTL_03],   " & _
            "'' [PROZT_03],  " & _
            "'' [BURKS_04], " & _
            "'' [KOSTL_04],  " & _
            "'' [PROZT_04], " & _
            "'' [BURKS_05], " & _
            "'' [KOSTL_05],  " & _
            "'' [PROZT_05], " & _
            "'' [BURKS_06], " & _
            "'' [KOSTL_06],  " & _
            "'' [PROZT_06],  " & _
            "'' [BURKS_07], " & _
            "'' [KOSTL_07],  " & _
            "'' [PROZT_07],  " & _
            "'' [BURKS_08], " & _
            "'' [KOSTL_08],  " & _
            "'' [PROZT_08],  " & _
            "'' [BURKS_09], " & _
            "'' [KOSTL_09],  " & _
            "'' [PROZT_09],  " & _
            "'' [BURKS_10], " & _
            "' ' [KOSTL_10],  " & _
            "' ' [PROZT_10]," & _
            "' ' [BURKS_11], " & _
            "' ' [KOSTL_11],  " & _
            "' ' [PROZT_11] " & _
            "FROM  dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
            "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
            "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
            "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1   " & _
            "WHERE  (P1.CO_EMPR = '10') " & _
            "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
            "AND P1.TI_SITU = 'ACT' " & _
            "ORDER BY  MT.IDUSR   "

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 2.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 3.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)

        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i

                    Case 1
                        'Validamos que no sea vacio 
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If
                    Case 2
                        'Validamos el campo nulo 
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                        End If
                    Case 5
                        'Validamos el campo nulo 
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                        End If
                    Case 8
                        'Validamos el campo nulo 
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                        End If

                    Case 9
                        'Validamos el campo nulo 
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next
            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()
        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()
        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '------------------------------------------------------------------------------
    '  Rutina de carga de datos maestro de Seguros 
    '------------------------------------------------------------------------------
    Sub f_pait0037()
        On Error GoTo ErrorHandler

        ' 1.-Devuelve el objeto Connection
        o_conex = cl_PA.conexiondb()
        ' 2.-Query de la consulta 
        p_ssql = " SELECT  MT.IDUSR [PERNR]," & _
                     "(SELECT  TOP 1 FE_USUA_CREA   FROM dbo.TDESTA_TRAB AS ET  " & _
                     "WHERE (CO_TRAB =  N1.CO_TRAB)  " & _
                     "AND (NU_ANNO = SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),103),7,4))  " & _
                     "AND CO_ESTA IN ('SCRTRT','VIDALE'))   [BEGDA] , " & _
                     "''   [ENDDA], " & _
                     "(SELECT  TOP 1 CO_ESTA  FROM dbo.TDESTA_TRAB AS ET  " & _
                     "WHERE (CO_TRAB =  N1.CO_TRAB)  " & _
                     "AND (NU_ANNO = SUBSTRING(CONVERT(VARCHAR(10),GETDATE(),103),7,4))  " & _
                     "AND CO_ESTA IN ('SCRTRT','VIDALE'))  [SUBTY], " & _
                     "''   [VSGES], " & _
                     "''   [VSNUM], " & _
                     "'1' [VSTAR]  " & _
                     "FROM  dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
                     "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
                     "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
                     "dbo.MTRA  AS MT  ON P1.CO_TRAB = MT.STCD1 " & _
                     "WHERE P1.CO_EMPR = '10' " & _
                     "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
                     "ORDER BY MT.IDUSR"

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 3.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 4.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)

        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i
                    Case 3
                        'Valida los nulos
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()
                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub

ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '------------------------------------------------------------------------------
    ''-----------------------------------------------------------------------------
    ''Rutina para la carga de datos maestros de Préstamos Cuota fija
    ''-----------------------------------------------------------------------------
    Sub f_pait0045()
        On Error GoTo ErrorHandler

        ' 1.-Devuelve el objeto Connection
        o_conex = cl_PA.conexiondb()
        ' 2.-Query de la consulta 
        p_ssql = "SELECT " & _
                     "MT.IDUSR [PERNR], " & _
                     "SP.TI_SOLI_PRES  [SUBTY],   " & _
                     "SP.FE_SOLI_PRES  [BEGDA], " & _
                     "' ' [ENDDA],  " & _
                     "SP.FE_APRO_PRES  [DATBW],  " & _
                     "SP.IM_SOLI_PRES   [DARBT],  " & _
                     "SP.CO_MONE          [DBTCU], " & _
                     "'01'                          [DKOND], " & _
                     "' '                             [TILBG], " & _
                     "SP.IM_AMOR_PRES   [TILBT], " & _
                     "SP.CO_MONE           [TILCU], " & _
                     "''  [ZAHLD01],  " & _
                     "'0150'  [ZAHLA01] " & _
                    "FROM  dbo.TMTRAB_EMPR AS P1 INNER JOIN  " & _
                    "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB  INNER JOIN  " & _
                    "dbo.TMSOLI_PRES AS SP ON SP.CO_TRAB = N1.CO_TRAB INNER JOIN  " & _
                    "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN  " & _
                    "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1   " & _
                    "WHERE P1.CO_EMPR = '10'  " & _
                    "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011')  " & _
                    "AND SP.TI_SITU = 'APR'   " & _
                    "AND P1.TI_SITU = 'ACT'   " & _
                    "ORDER BY MT.IDUSR  "

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 3.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 4.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)
        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i
                    Case 1
                        'Validamos el campo nulo
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value

                            ''validamos que en la cadena de texto contenga una ','
                            'If w_valor.Contains(",") = True Then
                            '    'Reemplasamos la coma que hay en una cadena de texto por espacio
                            '    w_valor = Replace(w_valor, ",", " ")
                            'End If

                        End If
                    Case 2
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If
                    Case 4
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If
                    Case 6
                        'Retorna la equivalencia moneda
                        w_valor = cl_PA.equivalenciaMoneda(o_record_sql_sel.Fields(c).value)
                    Case 10
                        'Retorna la equivalencia moneda
                        w_valor = cl_PA.equivalenciaMoneda(o_record_sql_sel.Fields(c).value)

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub

ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '"----------------------------------------------------------------------------
    'Rutina para la carga de datos maestros de prestamos cuota  
    'nomina especial 
    '"----------------------------------------------------------------------------
    Sub f_pait0045_1()

        On Error GoTo ErrorHandler

        ' 1.-Devuelve el objeto Connection
        o_conex = cl_PA.conexiondb()
        ' 2.-Query de la consulta 
        p_ssql = "SELECT " & _
                   "MT.IDUSR  [PERNR], " & _
                   "SP.TI_SOLI_PRES   [SUBTY],   " & _
                   "SP.FE_SOLI_PRES  [BEGDA], " & _
                   "' ' [ENDDA],  " & _
                   "SP.FE_APRO_PRES   [DATBW], " & _
                   "SP.IM_SOLI_PRES     [DARBT],  " & _
                   "SP.CO_MONE            [DBTCU],  " & _
                   "'01'                            [DKOND],  " & _
                   "' '                               [TILBG],    " & _
                   "SP.IM_AMOR_PRES   [TILBT], " & _
                   "SP.CO_MONE            [TILCU], " & _
                   "SP.NU_CUOT_SOLI   [ZAHLD01], " & _
                   "'0150'  [ZAHLA01] , " & _
                   "''  [ZAHLD02], " & _
                   "''  [ZAHLA02], " & _
                   "SP.IM_AMOR_PRES  [BETRG02], " & _
                   "''  [ZAHLD03], " & _
                   "''  [ZAHLA03], " & _
                   "SP.IM_AMOR_PRES  [BETRG03], " & _
                   "''  [ZAHLD04], " & _
                   "''  [ZAHLA04], " & _
                   "SP.IM_AMOR_PRES   [BETRG04], " & _
                   "''  [ZAHLD05], " & _
                   "''  [ZAHLA05], " & _
                   "SP.IM_AMOR_PRES   [BETRG05], " & _
                   "''  [ZAHLD06], " & _
                   "''  [ZAHLA06], " & _
                   "SP.IM_AMOR_PRES  [BETRG06], " & _
                   "''  [ZAHLD07], " & _
                   "''  [ZAHLA07], " & _
                   "SP.IM_AMOR_PRES  [BETRG07]  " & _
                   "FROM  dbo.TMTRAB_EMPR AS P1 INNER JOIN  " & _
                  "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB  INNER JOIN  " & _
                  "dbo.TMSOLI_PRES AS SP ON SP.CO_TRAB = N1.CO_TRAB INNER JOIN  " & _
                  "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN  " & _
                  "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1   " & _
                  "WHERE P1.CO_EMPR = '10'  " & _
                  "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011')  " & _
                  "AND SP.TI_SITU = 'APR'   " & _
                  "AND P1.TI_SITU = 'ACT'   " & _
                  "ORDER BY MT.IDUSR"
        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 3.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 4.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)
        'Sumatoria de cuotas
        Dim w_scuot As Double
        'Importe de prestamo
        Dim w_imp As Double
        'Cuota de prestamos
        Dim w_cuot As String
        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""
            w_scuot = 0
            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i
                    Case 1
                        'Validamos el campo nulo
                        If o_record_sql_sel.Fields(c).value Is DBNull.Value = True Then
                            w_valor = String.Empty
                        Else
                            w_valor = o_record_sql_sel.fields(c).value
                            ''validamos que en la cadena de texto contenga una ','
                            'If w_valor.Contains(",") = True Then
                            '    'Reemplasamos la coma que hay en una cadena de texto por espacio
                            '    w_valor = Replace(w_valor, ",", " ")
                            'End If

                        End If
                    Case 2
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If
                    Case 4
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case 6
                        'Retorna la equivalencia moneda
                        w_valor = cl_PA.equivalenciaMoneda(o_record_sql_sel.Fields(c).value)
                    Case 10
                        'Seteamos a PEN 
                        w_valor = cl_PA.equivalenciaMoneda(o_record_sql_sel.Fields(c).value)
                    Case 15
                        w_imp = Convert.ToDouble(o_record_sql_sel.Fields(5).value().ToString)
                        w_cuot = o_record_sql_sel.Fields(9).value.ToString()
                        If w_cuot > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(w_cuot) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            w_valor = ""
                        End If

                    Case 18
                        w_cuot = o_record_sql_sel.Fields(9).value.ToString()
                        If w_cuot > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(o_record_sql_sel.Fields(9).value.ToString) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            w_valor = ""
                        End If
                    Case 21
                        w_cuot = o_record_sql_sel.Fields(9).value.ToString()
                        If w_cuot > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(o_record_sql_sel.Fields(9).value.ToString) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            w_valor = ""
                        End If

                    Case 24
                        w_cuot = o_record_sql_sel.Fields(9).value.ToString()
                        If w_cuot > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(o_record_sql_sel.Fields(9).value.ToString) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            w_valor = ""
                        End If

                    Case 27
                        'Importe de la cuota
                        w_cuot = o_record_sql_sel.Fields(9).value.ToString()
                        'Validamos cuota > 0
                        If w_cuot > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(o_record_sql_sel.Fields(9).value.ToString) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            'No tiene importe cuota a mostrar
                            w_valor = ""
                        End If

                    Case 30
                        w_imp = Convert.ToDouble(o_record_sql_sel.Fields(5).value().ToString)
                        'cuota > 0 
                        If o_record_sql_sel.Fields(9).value.ToString() > 0 Then
                            'Calcula la sumatoria de cuotas
                            w_scuot = Convert.ToDouble(o_record_sql_sel.Fields(9).value.ToString) + w_scuot

                            If (w_scuot <= w_imp) Then
                                w_valor = o_record_sql_sel.Fields(c).value
                            Else
                                'No tiene importe cuota a mostrar
                                w_valor = ""
                            End If
                        Else
                            'No tiene importe cuota a mostrar
                            w_valor = ""
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1

            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub

ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '"-------------------------------------------------------------------
    ''Rutina de carga de maestro de Información de tiempos
    ''-------------------------------------------------------------------
    Sub f_pait0050()
        On Error GoTo ErrorHandler

        ' 1.- Retorna la cadena de conexion 
        o_conex = cl_PA.conexiondb()

        ' 2.-Query de la consulta 
        p_ssql = " SELECT " & _
                  "MT.IDUSR  [PERNR], " & _
                  "P1.FE_INGR_EMPR [BEGDA], " & _
                  "'' [ENDDA], " & _
                  "'' [ZAUSW], " & _
                  "'01' [PMBDE], " & _
                  "'001' [BDEGR] , " & _
                  "'001' [GRAWG],  " & _
                  "'001' [GRELG],  " & _
                  "FROM  dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
                  "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
                  "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
                  "dbo.vCATEGORIAS AS CT ON CT.CATEGORIA = P1.CO_PLAN INNER JOIN " & _
                  "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1  " & _
                  "WHERE (P1.CO_EMPR = '10') " & _
                  "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
                  "AND P1.TI_SITU = 'ACT' " & _
                  "ORDER BY  MT.IDUSR "
        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 4.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 5.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)

        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i

                    Case 1
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case 2
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()
    End Sub
    ''------------------------------------------------------------------------
    ''Rutina de carga de maestro de  Informacion de Beneficios
    ''-----------------------------------------------------------------------
    Sub f_pait0171()
        On Error GoTo ErrorHandler
        '1.- Retorna la conex a la base de datos 
        o_conex = cl_PA.conexiondb()

        p_ssql = " SELECT " & _
               "MT.IDUSR  [PERNR], " & _
               "P1.FE_INGR_EMPR [BEGDA], " & _
               "'' [ENDDA], " & _
               "'01' [BAREA], " & _
               "'0001' [BENGR], " & _
               "'0001' [BSTAT] " & _
               "FROM dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
               "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
               "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
               "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1 LEFT JOIN " & _
               "WHERE     (P1.CO_EMPR = '10') " & _
               "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
               "AND P1.TI_SITU = 'ACT' " & _
               "ORDER BY  MT.IDUSR "

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 4.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 5.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)


        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i

                    Case 1
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case 2
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    ''----------------------------------------------------------
    '' Rutina de carga de maestros de Planes de Salud 
    ''---------------------------------------------------------
    Sub f_pait0167()
        On Error GoTo ErrorHandler
        '1.- Retorna la conex a la base de datos 
        o_conex = cl_PA.conexiondb()

        p_ssql = " SELECT " & _
               "MT.IDUSR  [PERNR], " & _
               "P1.FE_INGR_EMPR [BEGDA], " & _
               "'' [ENDDA], " & _
               "'01' [BAREA], " & _
               "'0001' [BENGR], " & _
               "'0001' [BSTAT] " & _
               "FROM dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
               "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
               "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
               "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1 LEFT JOIN " & _
               "WHERE     (P1.CO_EMPR = '10') " & _
               "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
               "AND P1.TI_SITU = 'ACT' " & _
               "ORDER BY  MT.IDUSR "

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 4.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 5.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)


        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i

                    Case 1
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case 2
                        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
                            'Cambia de formato a dd/MM/yyy
                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                        Else
                            w_valor = o_record_sql_sel.Fields(c).value
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()

                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '"----------------------------------------------------------------------------------------
    '  Rutina para la carga de datos maestros de Contingente de Absentismos
    '-----------------------------------------------------------------------------------------
    Sub f_pait2006()

        '1.- Retorna la conex a la base de datos 
        Dim o_cadena As String
        o_cadena = My.Settings.cnNorthwind
        Dim o_conect As New SqlConnection(My.Settings.cnNorthwind)
        ' abre la conexión a la base de datos

        p_ssql = " SELECT " &
                 "O.ORDERID,  P.PRODUCTNAME FROM [Order Details] O INNER JOIN PRODUCTS P " &
                 "on P.PRODUCTID = O.ProductID "

        Dim o_adapter As New SqlDataAdapter()
        Dim o_edataset As New DataSet()
        Dim o_dtaview As DataView

        Try

            Dim o_command As New SqlCommand(p_ssql, o_conect)
            o_adapter.SelectCommand = o_command
            o_adapter.Fill(o_edataset, "Pedidos")
            o_adapter.Dispose()
            'Llenado de un dataview 
            o_dtaview = o_edataset.Tables("Pedidos").DefaultView
            Dim drv As DataRowView
            Dim s As String = ""

            'Recorro las filas del dataview 
            For Each drv In o_dtaview
                s &= drv("ORDERID").ToString & ","
                s &= drv("PRODUCTNAME").ToString & ControlChars.CrLf

                w_linea = w_linea & s
                s = ""
            Next

            'Ruta del fichero a descargar
            w_ruta = "F:\ZSQ_DOFIS_PY\txt\"

            'Nombre del fichero a descargar
            w_nfile = w_ruta & "" & PA & "" & ".txt"

            'Borra el fichero generado previamente
            If My.Computer.FileSystem.FileExists(w_nfile) Then
                My.Computer.FileSystem.DeleteFile(w_nfile)
            End If

            '' recorre las columnas, y genera el encabezado
            'For i = 0 To o_record_sql_sel.Fields.Count - 1
            '    w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
            'Next

            'Abre un fichero para escritura
            b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

            ''Escribe el detalle del fichero 
            b_filtxt.WriteLine(w_linea)

            'Cierra el fichero generado  
            b_filtxt.Flush()
            b_filtxt.Close()

            'Cierra el formulario una ves generado el fichero
            Me.Close()
            Me.Dispose()

            o_conex.Close()
            o_conex = Nothing

        Catch ex As Exception

        End Try



        ' 2.-Recupera parametros del proceso
        i = 1

        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next


        'Escribe en el fichero los errores generados 
        'b_filtxt.WriteLine(p_ssql)
        'b_filtxt.WriteLine(Err.Description)
        'b_filtxt.Close()
        ''Variable del personal act, ant
        'Dim w_idusr, w_idusra As String
        ''variable de fecha de inicio, fin valides
        'Dim w_fechfv As Date
        'Dim w_fechiv As String
        ''Variable de fecha de inicio de liq de contingentes , fin de liquidacion contingentes
        'Dim w_fechilc, w_fechflc As String
        ''Variable de fecha de vaciones
        'Dim w_fechvca As String
        'Dim w_fini_em As String
        ''Variable para calculo de fechas
        'Dim w_fech As String
        'Dim w_fch As Date
        ''Variable para calculo de diass
        'Dim w_dvaca As String
        'Dim w_d, w_dant, w_ddias, w_m As String
        ''Inicialiso las variables
        'w_dant = ""
        'w_fini_em = ""
        'w_fechilc = ""
        'w_fechflc = ""
        'w_idusra = ""
        'w_idusr = ""
        'w_fechvca = ""
        'w_d = ""

        '        ' recorre las filas y genera detalle
        '        Do While Not o_record_sql_sel.EOF
        '            ' recorre los campos en el registro actual del recordset para recuperar el dato
        '            w_linea = ""

        '            For i = 0 To o_record_sql_sel.Fields.Count - 1
        '                Select Case i
        '                    Case 1
        '                        'Guardamos la fecha inicio de vacaciones
        '                        w_fechvca = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
        '                        'Seteamos el valor a mostrar con 90 
        '                        w_valor = "90"
        '                    Case 2
        '                        'Fecha de inicio de validez
        '                        w_idusr = o_record_sql_sel.Fields(0).value
        '                        'Validamos que sea el mismo empleado 
        '                        If (w_idusra = w_idusr) Then
        '                            'Seteamos la fecha de inicio de valides igual a la fecha de inicio de liqu de conting
        '                            w_valor = w_fechilc
        '                        Else
        '                            'Cambia de formato a dd/MM/yyy
        '                            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
        '                            'Calculo del dia anterior 
        '                            w_d = Mid(w_valor, 1, 2)
        '                            If w_d > 1 Then
        '                                w_dant = Mid(w_valor, 1, 2) - 1
        '                            Else
        '                                w_dant = "01"
        '                            End If
        '                            'Guardo la fecha de ing a la empresa
        '                            w_fini_em = w_valor
        '                        End If

        '                    Case 3
        '                        'Fecha de fin de validez 
        '                        w_idusr = o_record_sql_sel.Fields(0).value
        '                        'Validamos que sea el mismo empleado 
        '                        If (w_idusra = w_idusr) Then
        '                            'Seteamos la fech fin de valides con el valor de fecha fin de liq de conting
        '                            w_valor = w_fechflc
        '                        Else
        '                            'Calculos para la fecha fin de valides 
        '                            w_fech = w_dant & "/" & Mid(w_fini_em, 4, 10)
        '                            w_fechfv = DateAdd(DateInterval.Year, 1, CDate(w_fech))
        '                            'Guardamos la fecha de fin de valides 
        '                            w_valor = CStr(w_fechfv)
        '                        End If

        '                    Case 4
        '                        'Fecha de inicio de liqu de conting
        '                        w_idusr = o_record_sql_sel.Fields(0).value
        '                        'Validamos que sea el mismo empleado 
        '                        If (w_idusra = w_idusr) Then
        '                            'Guardamos fecha de incio de liqu de contingentes
        '                            w_fechilc = DateAdd(DateInterval.Year, 1, CDate(w_fechilc))
        '                            'calculo para la fecha de inicio de lique de contingentes 
        '                            w_valor = w_fechilc
        '                        Else
        '                            'guardamos fecha de inicio de liqu de contingentes
        '                            w_fechilc = DateAdd(DateInterval.Year, 1, CDate(w_fini_em))
        '                            w_valor = w_fechilc
        '                        End If

        '                    Case 5
        '                        'guardamos fecha de fin de liqu de contingentes
        '                        w_fech = w_dant & "/" & Mid(w_fechilc, 4, 10)
        '                        w_fechflc = DateAdd(DateInterval.Year, 1, CDate(w_fech))
        '                        w_valor = CStr(w_fechflc)
        '                    Case 6
        '                        'Calculo de la cant de contingentes de tiempo de personal
        '                        w_dvaca = o_record_sql_sel.Fields(c).value()
        '                        'Validamos que sea el mismo empleado 
        '                        If (w_idusra = w_idusr) Then
        '                            'Seteamos la fecha de inicio de valides igual a la fecha de inicio de liqu de conting
        '                            w_fechiv = w_fechilc
        '                        Else
        '                            'Cambia de formato a dd/MM/yyy 
        '                            w_valor = Format(o_record_sql_sel.Fields(2).value, "dd/MM/yyyy")
        '                            'Guardo la fecha de inicio de la empresa
        '                            w_fechiv = w_valor
        '                        End If
        '                        Dim w_fchvca, w_fchflc As Date
        '                        w_fchvca = CStr(w_fechvca)
        '                        w_fchflc = CStr(w_fechflc)
        '                        'Si estamos dentro de mi rango de valides
        '                        If w_fchvca <= w_fechfv Then
        '                            w_valor = 0
        '                            'Validamos que nuestra fecha de vaciones a estado dentro del rango de liquidacion cerrado
        '                        ElseIf (w_fchvca <= w_fchflc) Then
        '                            If w_dvaca = 30 Then
        '                                'Blanquea registro si se a usado todas las vacaciones 
        '                                w_valor = ""
        '                            Else
        '                                'Guardamos la cant de conting de tiempo de personal
        '                                w_valor = 30 - w_dvaca
        '                            End If
        '                        Else

        '                        End If

        '                    Case Else
        '                        w_valor = o_record_sql_sel.Fields(c).value()

        '                End Select
        '                'Almacena el campo recuperado del recordset concatenado con el palote
        '                w_linea = w_linea & w_valor & ","
        '                'incrementa la columna actual   
        '                c = c + 1
        '                w_valor = ""
        '            Next
        '            'Escribe el detalle del txt 
        '            b_filtxt.WriteLine(w_linea)
        '            'Resetea el indice de las columnas
        '            c = 0
        '            'Referencia al registro actual (incrementa )
        '            f = f + 1
        '            w_idusra = w_idusr

        '            'Siguiente registro
        '            o_record_sql_sel.MoveNext()

        '        Loop

        '        ' cierra y descarga las referencias
        '        On Error Resume Next
        '        o_record_sql_sel.Close()

        '        o_record_sql_sel = Nothing

        '        'Cierra el fichero generado  
        '        b_filtxt.Flush()
        '        b_filtxt.Close()

        '        'Cierra el formulario una ves generado el fichero
        '        Me.Close()
        '        Me.Dispose()
        '        Exit Sub
        'ErrorHandler:
        '        'Escribe en el fichero los errores generados 
        '        b_filtxt.WriteLine(p_ssql)
        '        b_filtxt.WriteLine(Err.Description)
        '        b_filtxt.Close()

    End Sub
    '"---------------------------------------------------------------------------------------------------------
    '  Rutina para la carga de datos maestros de Contingente de Absentismos (vacaciones) 
    '---------------------------------------------------------------------------------------------------------
    Sub f_pait2006_1()
        On Error GoTo ErrorHandler
        '1.- Retorna la conex a la base de datos 
        o_conex = cl_PA.conexiondb()

        p_ssql = " SELECT " & _
                     " MT.IDUSR  [PERNR],  " & _
                     "'70' [KTART_01], " & _
                     "'20'  [ANZHL], " & _
                     "'01/01/2018'  [BEGDA_01], " & _
                     "'31/01/2018'  [ENDDA_01], " & _
                     "'01/01/2018'  [DESTA_01], " & _
                     "'31/01/2018'  [DEEND_01]  " & _
                       "FROM dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
                       "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
                       "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
                       "dbo.MTRA AS MT  ON P1.CO_TRAB = MT.STCD1  " & _
                       "WHERE     (P1.CO_EMPR = '10') " & _
                       "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
                       "AND P1.TI_SITU = 'ACT' " & _
                       "ORDER BY  MT.IDUSR "

        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 4.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 5.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)

        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)


        ' recorre las filas y genera detalle
        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1

                w_valor = o_record_sql_sel.Fields(c).value()

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1
                w_valor = ""
            Next

            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop

        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()

        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()

    End Sub
    '"----------------------------------------------------------------------------------------
    'Rutina para la carga de datos maestros de Saldos de tiempos 
    '-----------------------------------------------------------------------------------------
    Sub f_pait2012()
        On Error GoTo ErrorHandler

        ' 1.- Retorna la cadena de conexion 
        o_conex = cl_PA.conexiondb()
        ' 2.- Query de la consulta
        p_ssql = " SELECT " & _
                        " MT.IDUSR  [PERNR],  " & _
                        "''  [SUBTY], " & _
                        "P1.FE_INGR_EMPR [BEGDA], " & _
                        "''  [ENDDA], " & _
                        "P1.FE_INGR_EMPR  [ANZHL] " & _
                        "FROM dbo.TMTRAB_EMPR AS P1 INNER JOIN " & _
                        "dbo.TMTRAB_PERS AS N1 ON P1.CO_TRAB = N1.CO_TRAB INNER JOIN " & _
                        "dbo.TDCCOS_TRAB AS C1 ON P1.CO_TRAB = C1.CO_TRAB INNER JOIN " & _
                        "dbo.MTRA  AS MT  ON P1.CO_TRAB = MT.STCD1  " & _
                         "WHERE     (P1.CO_EMPR = '10') " & _
                        "AND (SUBSTRING(CAST(C1.CO_CENT_COST AS CHAR), 1, 3) = '011') " & _
                        "AND P1.TI_SITU = 'ACT' " & _
                         "ORDER BY  MT.IDUSR "
        'Ruta del fichero a descargar
        w_ruta = "D:\zprog_hr\ZSQ_DOFIS_PY\txt"
        'Nombre del fichero a descargar
        w_nfile = w_ruta & "" & PA & "" & ".txt"

        'Borra el fichero generado previamente
        If My.Computer.FileSystem.FileExists(w_nfile) Then
            My.Computer.FileSystem.DeleteFile(w_nfile)
        End If

        ' 3.-Recupera parametros del proceso
        i = 1
        For Each o_argument As String In My.Application.CommandLineArgs
            If i = 1 Then
                w_nfile = o_argument
            Else
                p_ssql = o_argument
                p_ssql = Replace(p_ssql, "#", "'")
            End If
            i = i + 1
        Next

        'Abre un fichero para escritura
        b_filtxt = My.Computer.FileSystem.OpenTextFileWriter(w_nfile, True)

        ' 4.-Seleccciona registros en  o_record_sql_sel
        ' abre la conexión a la base de datos
        o_conex.Open()

        ' crea un nuevo objeto recordset de seleccion de registros
        o_record_sql_sel = CreateObject("ADODB.Recordset")

        ' Ejecuta el sql para llenar el recordset
        o_record_sql_sel.Open(p_ssql, o_conex, 1, 3)
        ' variables para los indices de las filas y columnas
        c = 0
        f = 1

        w_linea = ""
        ' recorre las columnas, y genera el encabezado
        For i = 0 To o_record_sql_sel.Fields.Count - 1
            w_linea = w_linea & o_record_sql_sel.Fields(i).Name & ","
        Next
        'Escribe los encabezados del txt
        b_filtxt.WriteLine(w_linea)
        ' recorre las filas y genera detalle
        'variable fecha ingr a la empresa
        Dim w_fing As Date
        'variable para calculo de saldo de tiempos
        Dim w_fclc As Long

        Do While Not o_record_sql_sel.EOF
            ' recorre los campos en el registro actual del recordset para recuperar el dato
            w_linea = ""

            For i = 0 To o_record_sql_sel.Fields.Count - 1
                Select Case i
                    Case 2
                        'Guardo la fecha de ing a la empresa en inicio de la validez
                        w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
                    Case 3
                        'Formato de fecha del dia en dd/mm/yyy
                        If o_record_sql_sel.Fields(c).value.ToString = "" Then
                            w_valor = DateTime.Now.ToString("dd/MM/yyyy")
                        End If
                    Case 4
                        'Variable part ent , dec 
                        Dim w_ent, w_dec As String
                        'Guardamos la fecha de ingr de la empresa en formato dd/mm/yyyy 
                        w_fing = CDate(Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy"))
                        'Guardamos la dif de fechas 
                        w_fclc = DateDiff(DateInterval.Day, Now(), w_fing)
                        w_fclc = w_fclc * -1
                        If w_fclc >= 30 Then
                            'Guardamos el nro de meses
                            w_ent = w_fclc \ 30
                            'Guardamos el nro de dias restantes
                            w_dec = w_fclc Mod 30
                            'Guardamos el nro mes y dias
                            w_valor = w_ent & " m  " & w_dec & "d"
                        Else
                            w_valor = ""
                        End If

                    Case Else
                        w_valor = o_record_sql_sel.Fields(c).value()
                End Select

                'Almacena el campo recuperado del recordset concatenado con el palote
                w_linea = w_linea & w_valor & ","
                'incrementa la columna actual   
                c = c + 1

            Next
            'Escribe el detalle del txt 
            b_filtxt.WriteLine(w_linea)
            'Resetea el indice de las columnas
            c = 0
            'Referencia al registro actual (incrementa )
            f = f + 1
            'Siguiente registro
            o_record_sql_sel.MoveNext()

        Loop
        ' cierra y descarga las referencias
        On Error Resume Next
        o_record_sql_sel.Close()
        o_record_sql_ins.close()
        o_conex.Close()
        o_conex = Nothing
        o_record_sql_sel = Nothing
        o_record_sql_ins = Nothing

        'Cierra el fichero generado  
        b_filtxt.Flush()
        b_filtxt.Close()

        'Cierra el formulario una ves generado el fichero
        Me.Close()
        Me.Dispose()
        Exit Sub
ErrorHandler:
        'Escribe en el fichero los errores generados 
        b_filtxt.WriteLine(p_ssql)
        b_filtxt.WriteLine(Err.Description)
        b_filtxt.Close()
    End Sub
   
#End Region

End Class