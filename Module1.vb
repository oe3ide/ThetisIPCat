
Imports System.Text
Imports System.IO
Imports System.Net.Sockets
Imports System.Threading


Module Module1
    Private Declare Ansi Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Private Declare Ansi Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32

    Dim fName As String
    Dim errStr As String
    Dim IPStr As String = ""
    Dim PortStr As String = ""
    Dim port As Integer
    Dim delay As Int16
    Dim delayStr As String

    Dim tcpClient As New System.Net.Sockets.TcpClient()


    Dim returndata As String = ""

    Dim debug As String = ""
    Dim sFile As String


    Dim dir As String = Directory.GetCurrentDirectory()

    Sub Main()
        Dim startTime As Date = Now

        fName = dir + "\thetisipcat.ini"




        If File.Exists(fName) Then
            'read ini
            IPStr = INI_ReadValueFromFile("Network", "IP", "", fName)
            PortStr = INI_ReadValueFromFile("Network", "Port", "", fName)
            debug = INI_ReadValueFromFile("Common", "Debug", "", fName)
            delayStr = INI_ReadValueFromFile("Network", "DelayMS", "", fName)

            If delayStr = "" Then
                delay = 50
            Else
                Try
                    delay = CInt(delayStr)
                Catch ex As Exception
                    delay = 50
                End Try

            End If

            If Not isValidPort(PortStr) Then
                errStr = "PORT NOT VALID"
                noConnect(fName, IPStr, PortStr, errStr)
                Exit Sub
            End If

            If IPStr = "" Or PortStr = 0 Then
                MsgBox("No ini file found. Don't worry, we create now one :-)")
                writeINIDefaults(fName)
            End If
        Else
            'write defaults
            MsgBox("No ini file found. Don't worry, we create now one :-)")
            writeINIDefaults(fName)
        End If

        If Not IsIPAddressValid(IPStr) Then
            errStr = "IP-ADDRESS NOT VALID"
            noConnect(fName, IPStr, PortStr, errStr)
            Exit Sub
        End If



        Console.WriteLine("Using IP: " + IPStr)
        Console.WriteLine("Using Port: " + PortStr)
        Console.WriteLine("Using Delay (ms): " & delay)

        writeDebug("Using IP: " + IPStr)
        writeDebug("Using Port: " + PortStr)
        writeDebug("Using Delay (ms): " & delay)


        If My.Application.CommandLineArgs Is Nothing Then

            Console.WriteLine("No CAT Str")
            Exit Sub

        End If

        Dim cmds() As String
        Dim i As Integer


        i = 0
        For Each cmd In My.Application.CommandLineArgs
            cmd = cmd.ToUpper

            If InStr(cmd, "/") > 0 Then
                ReDim Preserve cmds(i)
                Dim tmp As String = Replace(cmd, "/", "")

                If InStr(tmp, "ZZ") > 0 Then

                    cmds(i) = tmp & ";"
                    i = i + 1
                End If
            End If

        Next

        Dim sendCat = Join(cmds, "")

        If sendCat = "" Then
            Console.WriteLine("Error: No valid arguments! Examples: /ZZFA /ZZSP1")
            Exit Sub
        End If

        Console.WriteLine("Sending CAT: " & sendCat)

        port = CInt(PortStr)

        If connect(IPStr, port) = False Then
            Console.WriteLine("Damn it! We have some serious trouble talking with Thetis!")
            Console.WriteLine("Error: " & errStr)

            writeDebug("Damn it! We have some serious trouble talking with Thetis!")
            writeDebug("Error: " & errStr)

            Exit Sub
        End If

        Dim networkStream As NetworkStream

        Try
            networkStream = tcpClient.GetStream()
        Catch ex As Exception
            errStr = ""
            noConnect(fName, IPStr, PortStr, errStr)
            Exit Sub
        End Try


        If networkStream.CanWrite And networkStream.CanRead Then

            returndata = writeTCP(tcpClient, networkStream, sendCat, True)
            tcpClient.Close()

        Else
            If Not networkStream.CanRead Then

                Console.WriteLine("cannot not write data to this stream")
                writeDebug("cannot not write data to this stream")
                tcpClient.Close()
            Else
                If Not networkStream.CanWrite Then

                    Console.WriteLine("cannot read data from this stream")
                    writeDebug("cannot read data from this stream")
                    tcpClient.Close()
                End If
            End If
        End If

        Dim runLength As Global.System.TimeSpan = Now.Subtract(startTime)
        Dim millisecs As Integer = runLength.Milliseconds


        Console.WriteLine("exec time (ms): " & millisecs)
        writeDebug("exec time (ms): " & millisecs)
        Console.WriteLine("returndata:" & returndata)

    End Sub



    Private Function writeTCP(TcpClient, NetworkStream, catStr, read)

        Try

            If Not NetworkStream.CanWrite Or Not NetworkStream.CanRead Then
                writeDebug("writeTCP called. catStr=" & catStr & " but Stream cannot read or write")
                connect(IPStr, port)
            End If

            Dim sendBytes As [Byte]() = Encoding.ASCII.GetBytes(catStr)

            writeDebug("writeTCP called. catStr=" & catStr)


            If read Then
                NetworkStream.Flush()
            End If

            NetworkStream.Write(sendBytes, 0, sendBytes.Length)

            If read Then

                Thread.Sleep(delay)
                Dim avail As Int16 = TcpClient.Available
                Dim bytes(avail) As Byte
                NetworkStream.Read(bytes, 0, avail)
                Return Encoding.ASCII.GetString(bytes, 0, avail)

            Else
                Exit Function
            End If

        Catch ex As Exception
            writeDebug("writeTCP exception:" & ex.Message)
        End Try

    End Function



    Private Function INI_ReadValueFromFile(ByVal strSection As String, ByVal strKey As String, ByVal strDefault As String, ByVal strFile As String) As String
        Dim strTemp As String = Space(1024), lLength As Integer
        lLength = GetPrivateProfileString(strSection, strKey, strDefault, strTemp, strTemp.Length, strFile)
        Return (strTemp.Substring(0, lLength))
    End Function

    Private Function INI_WriteValueToFile(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, ByVal strFile As String) As Boolean
        Return (Not (WritePrivateProfileString(strSection, strKey, strValue, strFile) = 0))
    End Function

    Private Sub writeINIDefaults(fName)

        INI_WriteValueToFile("Network", "IP", "127.0.0.1", fName)
        INI_WriteValueToFile("Network", "Port", "13013", fName)
        INI_WriteValueToFile("Network", "DelayMS", "50", fName)

        Process.Start(fName)

    End Sub

    Private Sub noConnect(fName, IPStr, PortStr, errStr)
        Console.WriteLine()

        If errStr = "" Then
            Console.WriteLine("Damn it! We couldn't connect to " & IPStr & " on port " & PortStr & ".")
            Console.WriteLine("1. check settings in Thetis (Setup->CAT Control->CAT+)")
            Console.WriteLine("2. check ini settings (press 'e' + <enter>)")
            Console.WriteLine("3. check firewall settings")

            writeDebug("Damn it! We couldn't connect to " & IPStr & " on port " & PortStr & ".")

        Else
            Console.WriteLine("Damn it! We have some serious trouble talking with Thetis!")
            Console.WriteLine("Error: " & errStr)

            writeDebug("Damn it! We have some serious trouble talking with Thetis!")
            writeDebug("Error: " & errStr)

        End If

        Console.WriteLine("")
        Console.WriteLine("<Press 'e' + <enter> to edit the ini-file or any other key to exit>")
        Dim inputStr As String = Console.Read()

        If inputStr = 101 Then

            Process.Start(fName)
            Exit Sub
        Else
            Exit Sub
        End If
    End Sub

    Private Function parseReturndata(returndata, parseStr, parseLength)

        If parseStr = "" Then
            Return ""
        End If

        If returndata = "" Or returndata Is Nothing Then
            Return ""
        End If

        returndata = returndata.Replace(Convert.ToChar(0), "")

        Dim werteS() As String
        werteS = returndata.split(";")
        For Each x In werteS
            Dim tmp As String = x
            tmp = x.Replace(vbLf, "")
            tmp = tmp.Replace(";", "")
            If tmp.IndexOf(parseStr) > -1 Then
                Return tmp.Substring(4, parseLength)

            End If
        Next



    End Function

    Public Function IsIPAddressValid(ByVal addrString As String) As Boolean

        Dim werte = addrString.Split(".")
        If werte.Length <> 4 Then
            Return False
        End If

        For Each y In werte

            If Not Integer.TryParse(y, vbNull) Then
                Return False
            End If

            If y > 255 Then
                Return False
            End If


        Next

        Return True

    End Function

    Function isValidPort(portStr) As Boolean

        If Not Double.TryParse(portStr, False) Then
            Return False
        Else
            Return True
        End If


    End Function

    Public socketexception As Exception
    Public TimeoutObject As New ManualResetEvent(False)
    Public IsConnectionSuccessful As Boolean
    Private Function connect(IPStr, port) As Boolean

        Try
            TimeoutObject.Reset()
            socketexception = Nothing
            tcpClient = New TcpClient
            tcpClient.SendTimeout = 5000
            tcpClient.ReceiveTimeout = 5000

            tcpClient.Client.BeginConnect(IPStr, port, New AsyncCallback(AddressOf CallBackMethod), tcpClient)
            If TimeoutObject.WaitOne(500, False) Then

                If IsConnectionSuccessful Then
                    Return True
                Else
                    'Throw socketexception
                    tcpClient.Close()
                    errStr = "Error connecting to Thetis"
                    Return False

                End If
            Else
                errStr = "Error connecting to Thetis"
                tcpClient.Close()
                Return False

            End If
        Catch ex As Exception

            Return False

        End Try


    End Function


    Private Sub CallBackMethod(ByVal asyncResult As IAsyncResult)

        Try
            IsConnectionSuccessful = False
            If Not tcpClient.Client Is Nothing Then
                Dim tcpCli As TcpClient = CType(asyncResult.AsyncState, TcpClient)
                If Not tcpCli Is Nothing Then
                    tcpCli.Client.EndConnect(asyncResult)
                    IsConnectionSuccessful = True
                End If
            End If

        Catch ex As Exception
            IsConnectionSuccessful = False
            socketexception = ex
        Finally
            TimeoutObject.Set()
        End Try
    End Sub


    Dim debugfile As String
    Dim fs As FileStream
    Public Sub writeDebug(txt As String)

        If debug <> "1" Then
            Exit Sub
        End If

        Dim txt2write As String

        If debugfile = "" Then
            debugfile = "thetisipcat" & System.DateTime.Now.ToString.Replace(":", "").Replace(" ", "-") & ".log"

            FileOpen(1, debugfile, OpenMode.Output, OpenAccess.Write, OpenShare.Shared)
        End If

        txt2write = System.DateTime.Now.ToString & ": " & txt & vbLf


        Print(1, txt2write)


    End Sub



End Module
