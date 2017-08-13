Imports System.Text.RegularExpressions
'*****************************************************************************
' Libreria de funciones de PA , OM
'****************************************************************************
Public Class c_PA
#Region "variables_pit"
    ' 1.-Define el objeto Connection
    Dim o_conect As New ADODB.Connection
    'Variable obj de conex
    Dim o_conex As Object
    Dim Conexion As String
    'Variables para la conex con la base de datos
    Dim p_server, p_basdat, p_usuari, p_passwr As String
    ' cursores para recuperar el rec de ado
    Dim o_record_sql_sel As Object
    Dim c, i As Integer
    Dim w_valor As String
#End Region

#Region "Funciones PA - OM"
    '********************************************************************
    ' Funcion que me retorna mi cadena de conexion
    '******************************************************************
    Function conexiondb()
        ' 1.-Define el objeto Connection
        o_conex = CreateObject("ADODB.Connection")

        ' 2.-Crea la cadena de conexión a usar
        p_server = "192.168.57.19"
        p_basdat = "OFIPLAN"
        p_usuari = "OFISIS"
        p_passwr = "CQ5001"

        Conexion = "Provider=SQLOLEDB.1;" & _
                   "Password=" & p_passwr & ";" & _
                   "Persist Security Info=True;" & _
                   "User ID=" & p_usuari & ";" & _
                   "Initial Catalog=" & p_basdat & ";" & _
                   "Data Source=" & p_server
        o_conex.ConnectionString = Conexion
        Return o_conex

    End Function
    '**********************************************************************
    ' Procedimiento que me va a permitir cambiar a fecha del dia 
    '**********************************************************************
    Sub formatoFechaHoy()
        If o_record_sql_sel.Fields(c).value.ToString <> " " Then
            'Cambia de formato a dd/MM/yyy
            w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
        Else
            w_valor = o_record_sql_sel.Fields(c).value
        End If
    End Sub
    '*********************************************************************
    'Funcion que me va a permitir extraer el valor numero de una direccion
    '*********************************************************************
    Public Function extraeNum(ByVal cadena As String)
        'Variable numerica de la direccion 
        Dim w_num As String
        w_num = ""
        w_valor = cadena
        If w_valor <> " " Then
            'Recorrer la cadena
            For j = 1 To w_valor.Length()
                'Evalua si la cadena es un numero
                If IsNumeric(Mid(cadena, j, 1)) Then
                    w_num = w_num & Mid(cadena, j, 1)
                End If

            Next j

        End If
        extraeNum = w_num
        w_valor = w_num
        Return w_valor

    End Function
    '************************************************************************************
    'Rutina equivalencia bancaria segun el codigo 
    '***********************************************************************************
    Public Function equivalenciaBanc(ByVal cadena As String, ByVal moneda As String) As String
        'Variable del tipo de mon de la cta bancaria
        Dim w_mon As String
        'Segun el cod del banco se muestra la equivalencia
        w_valor = cadena
        'Tipo de moneda de la cta bancaria
        w_mon = moneda
        Select Case w_valor
            Case "002"
                'BANCO DE CREDITO
                If w_mon = "SOL" Then
                    w_valor = "02P"
                Else
                    w_valor = "02S"
                End If
            Case "003"
                'INTERBANK 
                If w_mon = "SOL" Then
                    w_valor = "03S"
                Else
                    w_valor = "03D"
                End If
            Case "009"
                'BANCO SCOTIABANK S.A.A.
                If w_mon = "SOL" Then
                    w_valor = "09S"
                Else
                    w_valor = "09D"
                End If
            Case "011"
                'BANCO CONTINENTAL
                If w_mon = "SOL" Then
                    w_valor = "11S"
                Else
                    w_valor = "11D"
                End If
            Case "013"
                'CAJA TACNA
                If w_mon = "SOL" Then
                    w_valor = "A5S"
                Else
                    w_valor = "A5D"
                End If
            Case "018"
                'BANCO DE LA NACION
                If w_mon = "SOL" Then
                    w_valor = "18S"
                Else
                    w_valor = "18D"
                End If
            Case "038"
                'BANBIF
                If w_mon = "SOL" Then
                    w_valor = "38S"
                Else
                    w_valor = "38D"
                End If
            Case "056"
                'BANCO FALABELLA
                If w_mon = "SOL" Then
                    w_valor = "54S"
                Else
                    w_valor = "54D"
                End If

        End Select
        Return w_valor
    End Function
    '******************************************************************************
    'Rutina para setear de acuerdo al cod de zona la descripcion (Equivalencia)
    '******************************************************************************
    Public Function equivalenciaZona(ByVal cadena As String) As String
        'Obtenemos el codigo de la zona
        w_valor = cadena
        'Seteamos de acuerdo al cod de zona la descripcion (Equivalencia)
        Select Case w_valor
            Case "01"
                w_valor = "Urb. Urbanización"
            Case "02"
                w_valor = "P.J. pueblo joven"
            Case "03"
                w_valor = "U.V. unidad vecinal"
            Case "04"
                w_valor = "C.H. conjunto habitacional"
            Case "05"
                w_valor = "A.H. asentamiento humano"
            Case "06"
                w_valor = "Coo. Cooperativa"
            Case "07"
                w_valor = "Res. Residencial"
            Case "08"
                w_valor = "Z.I. zona industrial"
            Case "09"
                w_valor = "Gru. Grupo"
            Case "10"
                w_valor = "Cas. Caserío"
            Case "11"
                w_valor = "Fnd.Fundo"
            Case Else
                w_valor = "Otros"
        End Select
        Return w_valor
    End Function
    ''*****************************************************************************
    ''Rutina de cambio de formato a dd/mm/yyy 
    ''****************************************************************************
    'Public Sub formatofecha()

    '    If IsNothing(o_record_sql_sel.Fields(c).value) Then
    '        w_valor = ""
    '    End If
    '    If w_valor <> " " Then
    '        w_valor = Format(o_record_sql_sel.Fields(c).value, "dd/MM/yyyy")
    '    End If

    'End Sub
    '****************************************************************************
    ' Rutina de equivalencia segun el estado civil 
    '****************************************************************************
    Public Function equivalenciaEstadoCivil(ByVal cadena As String) As String
        'Seteamos segun el cod estado civil la descripcion (equivalencia) '
        w_valor = cadena
        Select Case w_valor
            Case "CAS"
                w_valor = "Casado"
            Case "SOL"
                w_valor = "Soltero"
            Case "VIU"
                w_valor = "Viudo"
            Case "DIV"
                w_valor = "Divorciado"
            Case "SEP"
                w_valor = "Separado"
            Case "UNI"
                w_valor = "Union de hecho"
            Case "CNV"
                w_valor = "Convivencia"
        End Select
        Return w_valor

    End Function
    '****************************************************************************
    ' Rutina para validacion de un correo electronico
    '****************************************************************************
    Public Function validaEmail(ByVal email As String) As Boolean
        Dim emailRegex As New System.Text.RegularExpressions.Regex(
            "^(?<user>[^@]+)@(?<host>.+)$")
        Dim emailMatch As System.Text.RegularExpressions.Match =
           emailRegex.Match(email)
        Return emailMatch.Success

    End Function
    '****************************************************************************
    ' Rutina para setear con  el cod del pais de acuerdo al nombre del pais 
    '****************************************************************************
    Public Function equivalenciaPais(ByVal cadena As String) As String
        'Setea el cod del pais de acuerdo al nombre del pais 
        w_valor = cadena
        Select Case w_valor
            Case "PERU"
                w_valor = "PER"
            Case "ARGENTINA"
                w_valor = "ARG"
            Case "BOLIVIA"
                w_valor = "BOL"
            Case "BRASIL"
                w_valor = "BRA"
            Case "ESTADOS UNIDOS"
                w_valor = "USA"
            Case "VENEZUELA"
                w_valor = "VEN"
            Case Else
                w_valor = ""
        End Select
        Return w_valor

    End Function
    '*******************************************************************
    ' Rutina de equivalencia del mes de nacimiento 
    '******************************************************************
    Public Function equivalenciaMesNacimiento(ByVal cadena As String) As String
        w_valor = cadena
        Select Case w_valor
            Case "01"
                w_valor = "Enero"
            Case "02"
                w_valor = "Febrero"
            Case "03"
                w_valor = "Marzo"
            Case "04"
                w_valor = "Abril"
            Case "05"
                w_valor = "Mayo"
            Case "06"
                w_valor = "Junio"
            Case "07"
                w_valor = "Julio"
            Case "08"
                w_valor = "Agosto"
            Case "09"
                w_valor = "Septiembre"
            Case "10"
                w_valor = "Octubre"
            Case "11"
                w_valor = "Noviembre"
            Case "12"
                w_valor = "Diciembre"
        End Select
        Return w_valor
    End Function
    '**********************************************************************
    ' Rutina de equivalencia segun el codigo de afp seteamos la descripcion 
    '**********************************************************************
    Public Function equivalenciaAfp(ByVal cadena As String) As String
        'Segun el cod del afp seteamos con su descripcion de acuerdo a la equivalencia
        w_valor = cadena
        Select Case w_valor
            'Seteamos la descripcion del AFP 
            Case "99"
                w_valor = "SRP"
            Case "HAB"
                w_valor = "HABITAT"
            Case "HOR"
                w_valor = "HORIZONTE"
            Case "INT"
                w_valor = "AFP INTEGRA"
            Case "NA"
                w_valor = "SNP"
            Case "PRI"
                w_valor = "PRIMA"
            Case "PRO"
                w_valor = "PROFUTURO"
            Case "NVU"
                w_valor = "UNION VIDA"
        End Select

        Return w_valor

    End Function
    '**********************************************************************
    ' Rutina de equivalencia segun el codigo de vivienda seteamos la descripcion 
    '**********************************************************************
    Public Function equivalenciaVivienda(ByVal cadena As String)
        'Seteamos la descripcion de la vivienda de acuerdo a la equivalencia
        w_valor = cadena
        Select Case w_valor
            Case "ALQ"
                w_valor = "Alquilado"
            Case "FAM"
                w_valor = "Familiar"
            Case "NA"
                w_valor = "No aplica"
            Case "PRO"
                w_valor = "Propia"

        End Select
        Return w_valor

    End Function
    '*************************************************************************
    ' Rutina de equivalencia segun el codigo de cod contrato seteamos el tipo
    ' de contrato 
    '************************************************************************
    Function equivalenciaTipoContrato(ByVal cadena As String)
        w_valor = cadena
        Select Case w_valor
            Case "IND"
                w_valor = "Indeterminado"
            Case "PAR"
                w_valor = "Tiempo parcial"
            Case "PLA"
                w_valor = "Incremento actividad"
            Case "NME"
                w_valor = "Necesidad mercado"
            Case "REM"
                w_valor = "Reconversión empres."
            Case "PRA"
                w_valor = "Prácticas pre profesionales"
            Case "PRO"
                w_valor = "Prácticas profesionales"
            Case "PRT"
                w_valor = "Tiempo parcial"
            Case "SUP"
                w_valor = "De suplencia"
            Case "EXT"
                w_valor = "De extranjero"
            Case Else
                w_valor = "Otros no previstos"
        End Select

        Return w_valor
    End Function
    '************************************************************************
    ' Rutina de equivalencia de desc. de la relacion laboral de acuerdo 
    ' codcalif del trabajador
    '***********************************************************************
    Function equivalenciaRelacionLaboral(ByVal cadena As String)
        'Seteamos la relacion laboral de acuerdo al codcalif del trabajador
        w_valor = cadena
        Select Case w_valor
            Case "001"
                w_valor = "Confianza fiscalizable"
            Case "002"
                w_valor = "Fiscalizable"
            Case "005"
                w_valor = "De dirección"
            Case "006"
                w_valor = "Confianza no fiscalizable"
            Case Else
                w_valor = "No aplica"
        End Select

        Return w_valor

    End Function
    '***********************************************************************
    '   Rutina de equivalencia de moneda
    '***********************************************************************
    Function equivalenciaMoneda(ByVal cadena As String)
        w_valor = cadena
        Select Case w_valor
            Case "SOL"
                w_valor = "PEN"
            Case "DOL"
                w_valor = "USD"
        End Select

        Return w_valor

    End Function

    '***********************************************************************
    '   Rutina de equivalencia el codigo del dcto de acuerdo al tipo de 
    '  documento
    '************************************************************************
    Function equivalenciaTipoDocumento(ByVal cadena As String)
        w_valor = cadena
        Select Case w_valor
            Case "DNI"
                w_valor = "1"
            Case "CEX"
                w_valor = "4"
            Case "RUC"
                w_valor = "6"
            Case "PAS"
                w_valor = "7"
            Case "PAR"
                w_valor = "11"
            Case "BRE"
                w_valor = "16"
        End Select
        Return w_valor

    End Function
    '***********************************************************************
    ' Rutina de equivalencia de Licencia de conducir
    '***********************************************************************
    Function equivalenciaLicenciaConducir(ByVal cadena As String)
        '0 w_valor 
        w_valor = cadena
        Select Case w_valor
            Case "A1"
                w_valor = "A-I"
            Case "A2"
                w_valor = "A-II a"
            Case "A2A"
                w_valor = "A-II a"
            Case "A2B"
                w_valor = "A-II b"
            Case "A3A"
                w_valor = "A-III a"
            Case "A3B"
                w_valor = "A-III b"
            Case "A3C"
                w_valor = "A-III c"

        End Select

        Return w_valor

    End Function
    Function equivalenciaTipoVivienda(ByVal cadena As String)
        w_valor = cadena
        Select Case w_valor
            Case "FAM"
                w_valor = "Familiar"
            Case "ALQ"
                w_valor = "Alquilada"
            Case "NA"
                w_valor = "No aplicable"
            Case "PRO"
                w_valor = "Propia"
        End Select
        Return w_valor

    End Function
    '**********************************************************************
    ' Rutina de equivalencia de acreditacion vinculacion familiar
    '**********************************************************************
    Function equivalenciaAcreeditacionFamiliar(ByVal cadena As String)
        w_valor = cadena
        Select Case w_valor
            Case "01"
                w_valor = "Escritura publica"
            Case "05"
                w_valor = "Acta o partida de matrimonio civil"
            Case "06"
                w_valor = "Acta o partida de matrimonio inscrito en reg consular peruano"
            Case "08"
                w_valor = "Escritura publica - reconoc. de union de hecho - ley n.29560 "
            Case "09"
                w_valor = "Resolucion judicial - reconoc de union de hecho "
            Case "10"
                w_valor = "Acta de nacimiento o documento analogo que sustenta filiacion"
        End Select

        Return w_valor

    End Function
    '************************************************************************
    ' Rutina de  redondeo a 2 decimales
    '***********************************************************************
    Function redondearNumero(ByVal Numero As String) As String
        Dim ParteEntera As String = Int(Numero)
        Dim ParteDecimal As String
        If Not (Len(Numero) - Len(ParteEntera)) = 0 Then
            ParteDecimal = Right(Numero, Len(Numero) - Len(ParteEntera) - 1)
        Else
            ParteDecimal = "00"
        End If
        Dim Num As Double
        If Len(ParteDecimal) >= 3 Then
            ParteDecimal = Left(ParteDecimal, 3)
            If Mid(ParteDecimal, 3, 1) >= "5" Then
                ParteDecimal = Left(ParteDecimal, 2)
                Num = Convert.ToDouble(ParteDecimal)
                Num = Num + 1
                If Len(CStr(Num)) = 3 Then ParteEntera = ParteEntera + 1
                ParteDecimal = Right(CStr(Num), 2)
            End If
        Else
            ParteDecimal = Left(ParteDecimal, 2)
        End If
        redondearNumero = ParteEntera & "." & ParteDecimal
        Return redondearNumero
    End Function
    '***********************************************************************
    ' Rutina que me devuelve el nro de dias de un determinado mes 
    '***********************************************************************
    Function DiasDelMes(ByVal fecha As String)
        Dim w_ndias As String
        Select Case Month(fecha)
            Case 2
                If Bisiesto(Year(fecha)) = True Then
                    w_ndias = 29
                Else
                    w_ndias = 28
                End If
            Case 1, 3, 5, 7, 8, 10, 12
                w_ndias = 31
            Case Else
                w_ndias = 30
        End Select

        DiasDelMes = w_ndias
    End Function

    Function Bisiesto(Num As Integer) As Boolean
        If Num Mod 4 = 0 And (Num Mod 100 Or Num Mod 400 = 0) Then
            Bisiesto = True
        Else
            Bisiesto = False
        End If
    End Function
#End Region

    '****************************************************************************
    ' Libreria de funciones para la conversion a csv desde un dataset 
    '****************************************************************************
    'Función de Conversión a CSV
    Private Function ConvierteCSV(ByRef pDatos As DataSet,
        Optional ByVal pDelimitadorColumnas As String = ",",
        Optional ByVal pDelimitadorRegistros As String = vbNewLine) As String

        'Variables Locales
        Dim mSalida As String
        Dim mRegistro As String
        Dim mContadorRegistros As Integer = 0
        Dim mContadorColumnas As Integer = 0
        Dim mValor As String = ""
        Dim mNombreColumna As String = ""

        'Solo si hay datos
        If Not IsNothing(pDatos) Then
            If pDatos.Tables.Count > 0 Then

                'fila de titulos de columna
                mContadorColumnas = 0
                mRegistro = ""
                'para cada columna
                While mContadorColumnas < pDatos.Tables(0).Columns.Count

                    mNombreColumna = pDatos.Tables(0).Columns(mContadorColumnas).ColumnName

                    If mRegistro <> "" Then
                        mRegistro = mRegistro & pDelimitadorColumnas
                    End If

                    mValor = """" & mNombreColumna & """"
                    mRegistro = mRegistro & mValor

                    mContadorColumnas = mContadorColumnas + 1

                End While
                mSalida = mSalida & mRegistro

                'procesa los registros del dataset
                While mContadorRegistros < pDatos.Tables(0).Rows.Count
                    If mSalida <> "" Then
                        mSalida = mSalida & pDelimitadorRegistros
                    End If
                    mRegistro = ""
                    mContadorColumnas = 0
                    'para cada columna	
                    While mContadorColumnas < pDatos.Tables(0).Columns.Count
                        If mRegistro <> "" Then
                            mRegistro = mRegistro & pDelimitadorColumnas
                        End If

                        mValor = pDatos.Tables(0).Rows(mContadorRegistros)(mContadorColumnas)
                        If InStr(mValor, pDelimitadorColumnas) > 0 Then
                            mValor = """" & mValor & """"
                        End If
                        mRegistro = mRegistro & mValor
                        mContadorColumnas = mContadorColumnas + 1
                    End While
                    'añade el registro
                    mSalida = mSalida & mRegistro
                    mContadorRegistros = mContadorRegistros + 1
                End While
            End If
        End If

        Return mSalida
    End Function

    Private Sub DescargaCSV(ByVal pCSV As String, ByVal pNombreCSV As String)

        '    'Obtiene la respuesta actual
        '    Dim response As Web.HttpResponse = System.Web.HttpContext.Current.Response

        '    'Borra la respuesta
        '    response.Clear()
        '    response.ClearContent()
        '    response.ClearHeaders()

        '    'Tipo de contenido para forzar la descarga
        '    response.ContentType = "application/octet-stream"
        '    response.AddHeader("Content-Disposition", "attachment; filename=" & pNombreCSV)

        '    'Convierte el string a array de bytes
        '    Dim buffer(Len(pCSV)) As Byte
        '    Dim mContador As Long = 0
        '    While mContador < Len(pCSV)
        '        buffer(mContador) = Asc(Mid(pCSV, mContador + 1, 1))
        '        mContador = mContador + 1
        '    End While

        '    'Envia los bytes
        '    response.BinaryWrite(buffer)
        '    response.End()

    End Sub



End Class
