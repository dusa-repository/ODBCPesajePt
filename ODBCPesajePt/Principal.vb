Imports System.Xml
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data
Imports System.Net
Module Principal

    Sub Main()

        'PLASPIDATA  LIBRERIA PARA SPI

        cargar_parametros()
        'importarOrdenesProduccion()
        importarLocalidades()
    End Sub

    Public conexionString As String = ""
    Dim lineaLogger As String
    Dim logger As StreamWriter
    Dim lineaLoggerE As String
    Dim Conn400 As ADODB.Connection
    Dim Rst400 As ADODB.Recordset

    Dim RstSQLAS As ADODB.Recordset
    Dim CmdSQLAS As ADODB.Command
    Dim ConnSQLAS As ADODB.Connection

    Dim cadenaSQL As String
    Dim cadenaAS400_DTA As String
    Dim cadenaAS400_CTL As String
    Dim unaVez As Boolean
    Dim server As String
    Dim database As String
    Dim uid As String
    Dim pwd As String


    Private Function JulianToSerial(ByVal JulianDate As Long) As Date


        Dim SerialDate As Date

        'Convert the Julian date to a serial date.
        SerialDate = DateSerial(1900 + Int(JulianDate / 1000), 1, _
           JulianDate Mod 1000)

        Return SerialDate
    End Function


    Private Function existeOrden(ByVal wadoco As String, ByVal walitm As String, ByVal watrdj As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select * from orden_produccion where numero='" + wadoco + "' and walitm='" + walitm + "' and fecha= convert(datetime,'" & Format(JulianToSerial(watrdj), "dd/MM/yyyy") & "',103)  "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try



        Return existe
    End Function

    Private Function existeItem(ByVal walitm As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select * from producto where codigo='" + walitm + "'  "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try



        Return existe
    End Function





    Private Function existeLocalidad(ByVal lmlocn As String) As Boolean

        Dim existe As Boolean
        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()

            cmdSQL.Connection = connSQL
            cmdSQL.CommandText = "select * from almacen where codigo='" + lmlocn + "' "
            Dim lrdSQL As SqlDataReader = cmdSQL.ExecuteReader()
            existe = False
            While lrdSQL.Read()
                existe = True
            End While

            lrdSQL.Close()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try



        Return existe
    End Function


    Private Sub importarOrdenesProduccion()


        Try

            'cadenaAS400_DTA = "DSN=SPI;uid=TRANSFTP;pwd=TRANSFTP;"
            Dim cnn400 As New Odbc.OdbcConnection(cadenaAS400_DTA)
            Dim rs400 As New Odbc.OdbcCommand("SELECT wasrst,watrdj,wadcto,wadoco,wammcu,walocn,walotn,wawr01,wawr02,wawr03,walitm,wadl01, wauom,wauorg,wasoqs FROM f4801 WHERE WAMCU ='    300A0051' and wasrst<'60'", cnn400)
            Dim reader400 As Odbc.OdbcDataReader

            cnn400.Open()
            reader400 = rs400.ExecuteReader


            Dim connSQL As New SqlConnection
            Dim cmdSQL As New SqlCommand

            Try

                connSQL.ConnectionString = conexionString
                connSQL.Open()

                cmdSQL.Connection = connSQL

                cmdSQL.CommandText = "UPDATE [dbo].[orden_produccion]   SET [estado] = 'Inactiva' "
                cmdSQL.ExecuteNonQuery()

                connSQL.Close()

            Catch ex As Exception

            End Try


            While reader400.Read()

                If existeItem(Trim(reader400("walitm"))) Then
                    If existeOrden(Trim(reader400("wadoco")), Trim(reader400("walitm")), Trim(reader400("watrdj"))) Then
                        guardarOrdenes(True, Trim(reader400("wasrst")), Trim(reader400("watrdj")), Trim(reader400("wadcto")), Trim(reader400("wadoco")), Trim(reader400("wammcu")), Trim(reader400("walocn")), Trim(reader400("walotn")), Trim(reader400("wawr01")), Trim(reader400("wawr02")), Trim(reader400("wawr03")), Trim(reader400("walitm")), Trim(reader400("wadl01")), Trim(reader400("wauom")), Trim(reader400("wauorg")), Trim(reader400("wasoqs")))
                    Else
                        guardarOrdenes(False, Trim(reader400("wasrst")), Trim(reader400("watrdj")), Trim(reader400("wadcto")), Trim(reader400("wadoco")), Trim(reader400("wammcu")), Trim(reader400("walocn")), Trim(reader400("walotn")), Trim(reader400("wawr01")), Trim(reader400("wawr02")), Trim(reader400("wawr03")), Trim(reader400("walitm")), Trim(reader400("wadl01")), Trim(reader400("wauom")), Trim(reader400("wauorg")), Trim(reader400("wasoqs")))
                    End If

                End If

            End While

            reader400.Close()
            cnn400.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try



    End Sub


    Private Sub guardarOrdenes(ByVal existe As Boolean, ByVal wasrst As String, ByVal watrdj As String, ByVal wadcto As String, ByVal wadoco As String, ByVal wammcu As String, ByVal walocn As String, ByVal walotn As String, ByVal wawr01 As String, ByVal wawr02 As String, ByVal wawr03 As String, ByVal walitm As String, ByVal wadl01 As String, ByVal wauom As String, ByVal wauorg As String, ByVal wasoqs As String)

        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()
            cmdSQL.Connection = connSQL
            If existe Then
                cmdSQL.CommandText = "UPDATE [dbo].[orden_produccion]   SET [estado] = 'Activa'     ,[fecha_auditoria] = convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103)     ,[hora_auditoria] = '" & Format(Now(), "hh:mm:ss") & "'     ,[usuario_auditoria] = 'interfaz'      ,[wadcto] = '" & wadcto & "'      ,[wammcu] = '" & wammcu & "'     ,[walocn] = '" & walocn & "'     ,[walotn] = '" & walotn & "'     ,[wawr01] = '" & wawr01 & "'      ,[wawr02] = '" & wawr02 & "'     ,[wawr03] = '" & wawr03 & "'         ,[wadl01] = '" & wadl01 & "'     ,[wauom] = '" & wauom & "'      ,[wauorg] = '" & CDbl(wauorg) / 1000 & "'     ,[wasoqs] = '" & CDbl(wasoqs) / 1000 & "' WHERE [fecha] = convert(datetime,'" & Format(JulianToSerial(watrdj), "dd/MM/yyyy") & "',103) AND [walitm] = '" & walitm & "'  AND [numero] = '" & wadoco & "'  "
            Else
                cmdSQL.CommandText = "INSERT INTO [dbo].[orden_produccion]([estado],[fecha],[fecha_auditoria],[hora_auditoria],[numero],[usuario_auditoria],[wadcto],[wammcu],[walocn],[walotn],[wawr01],[wawr02],[wawr03],[walitm],[wadl01],[wauom],[wauorg],[wasoqs])     VALUES('Activa',convert(datetime,'" & Format(JulianToSerial(watrdj), "dd/MM/yyyy") & "',103),convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" & Format(Now(), "hh:mm:ss") & "' ,'" & wadoco & "','interfaz','" & wadcto & "','" & wammcu & "','" & walocn & "','" & walotn & "','" & wawr01 & "','" & wawr02 & "','" & wawr03 & "','" & walitm & "','" & wadl01 & "','" & wauom & "','" & CDbl(wauorg) / 1000 & "','" & CDbl(wasoqs) / 1000 & "'   )"
            End If

            cmdSQL.ExecuteNonQuery()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try

    End Sub



    Private Sub importarLocalidades()

        Try
            'cadenaAS400_DTA = "DSN=SPI;uid=TRANSFTP;pwd=TRANSFTP;"
            Dim cnn400 As New Odbc.OdbcConnection(cadenaAS400_DTA)
            Dim rs400 As New Odbc.OdbcCommand("SELECT LMLOCN FROM f4100 WHERE LMMCU='    300A0005' AND LMLOCN<>''  ", cnn400)
            Dim reader400 As Odbc.OdbcDataReader

            cnn400.Open()
            reader400 = rs400.ExecuteReader


            Dim connSQL As New SqlConnection
            Dim cmdSQL As New SqlCommand

            Try

            Catch ex As Exception

            End Try


            While reader400.Read()

                Console.WriteLine(reader400("LMLOCN"))

                If existeLocalidad(Trim(reader400("LMLOCN"))) Then
                    guardarLocalidad(True, Trim(reader400("LMLOCN")))
                Else
                    guardarLocalidad(False, Trim(reader400("LMLOCN")))
                End If

            End While

            reader400.Close()
            cnn400.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try



    End Sub


    Private Sub guardarLocalidad(ByVal existe As Boolean, ByVal lmlocn As String)

        Dim connSQL As New SqlConnection
        Dim cmdSQL As New SqlCommand

        Try

            connSQL.ConnectionString = conexionString
            connSQL.Open()
            cmdSQL.Connection = connSQL
            If existe Then
                cmdSQL.CommandText = "UPDATE [dbo].[almacen]   SET [descripcion] = '" & lmlocn & "'    ,[fecha_auditoria] = convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103)     ,[hora_auditoria] = '" & Format(Now(), "hh:mm:ss") & "'     ,[usuario_auditoria] = 'interfaz'      WHERE [codigo] = '" & lmlocn & "'  "
            Else
                cmdSQL.CommandText = "INSERT INTO [dbo].[almacen]([fecha_auditoria],[hora_auditoria],[usuario_auditoria],[codigo],[descripcion])     VALUES(convert(datetime,'" & Format(Now(), "dd/MM/yyyy") & "',103),'" & Format(Now(), "hh:mm:ss") & "' ,'interfaz','" & lmlocn & "','" & lmlocn & "')"
            End If

            cmdSQL.ExecuteNonQuery()
            connSQL.Close()

        Catch ex As Exception
            escribirLog(ex.StackTrace.ToString & "-" & ex.Message.ToString, " ")
        End Try

    End Sub





    Private Sub cargar_parametros()

        Dim host As String
        Dim database As String
        Dim user As String
        Dim password As String

        Dim conexion As New SqlConnection
        Dim comando As New SqlClient.SqlCommand

        Dim diccionario As New Dictionary(Of String, String)
        Dim xmldoc As New XmlDataDocument()
        Dim file_log_path As String

        Try
            file_log_path = Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "")
            If System.IO.File.Exists(file_log_path & "\log.log") Then
            Else
                Dim fs1 As FileStream = File.Create(file_log_path & "\log.log")
                fs1.Close()
            End If

            Try
                logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.log", True)
            Catch ex As Exception

            End Try


            Dim fs As New FileStream(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\configuracion.xml", FileMode.Open, FileAccess.Read)
            xmldoc = New XmlDataDocument()
            xmldoc.Load(fs)
            diccionario = obtenerNodosHijosDePadre("parametros", xmldoc)
            host = diccionario.Item("host")
            database = diccionario.Item("database")
            user = diccionario.Item("user")
            password = diccionario.Item("password")
            cadenaAS400_DTA = diccionario.Item("DSN1")
            conexionString = "Data Source=" & host & ";Database=" & database & ";User ID=" & user & ";Password=" & password & ";"

        Catch oe As Exception
            escribirLog(oe.StackTrace.ToString & "-" & oe.Message.ToString, " ")
        Finally

            logger.Close()
        End Try

    End Sub


    Public Sub escribirLog(ByVal mensaje As String, ByVal proceso As String)

        Dim time As DateTime = DateTime.Now
        Dim format As String = "dd/MM/yyyy HH:mm "

        Try
            lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
            logger.WriteLine(lineaLogger)
            logger.Flush()

        Catch ex As Exception

            Try
                logger = New StreamWriter(Replace(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase), "file:\", "") & "\log.log", True)
                lineaLogger = proceso & time.ToString(format) & ":" & mensaje & vbNewLine
                logger.WriteLine(lineaLogger)
                logger.Flush()
            Catch ex1 As Exception


            End Try

        End Try


    End Sub

    Public Function obtenerNodosHijosDePadre(ByVal nombreNodoPadre As String, ByVal xmldoc As XmlDataDocument) As Dictionary(Of String, String)
        Dim diccionario As New Dictionary(Of String, String)
        Dim nodoPadre As XmlNodeList
        Dim i As Integer
        Dim h As Integer
        nodoPadre = xmldoc.GetElementsByTagName(nombreNodoPadre)
        For i = 0 To nodoPadre.Count - 1
            For h = 0 To nodoPadre(i).ChildNodes.Count - 1
                If Not diccionario.ContainsKey(nodoPadre(i).ChildNodes.Item(h).Name.Trim()) Then
                    diccionario.Add(nodoPadre(i).ChildNodes.Item(h).Name.Trim(), nodoPadre(i).ChildNodes.Item(h).InnerText.Trim())
                End If
            Next
        Next
        Return diccionario
    End Function


End Module
