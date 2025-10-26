Imports System.IO, System.Reflection
Imports System.Drawing, System.Windows.Forms
Public Class EmbeddedResource
    Public Shared ReadOnly Property AssemblyName() As String
        Get
            Return Assembly.GetEntryAssembly.GetName.Name.Replace(Space(1), "_")
        End Get
    End Property
    Shared Sub ListResourceNames(ByRef List As List(Of String))
        List.AddRange(Assembly.GetEntryAssembly.GetManifestResourceNames)
    End Sub
    Shared Function Exists(name As String) As Boolean
        Dim ResName As String = String.Format("{0}.{1}", AssemblyName, name)
        'Dim Index As Integer = Array.BinarySearch(Assembly.GetEntryAssembly.GetManifestResourceNames, ResName)
        'Return CBool(Index >= 0)
        For Each Item As String In Assembly.GetEntryAssembly.GetManifestResourceNames
            If Item = ResName Then
                Return True
            End If
        Next
        Return False
    End Function

    Shared Function GetStreamReader(Name As String) As StreamReader
        If Exists(Name) = False Then Throw New Exception("Embedded Resource Item, Not Found!")
        Dim SreamName As String = String.Format("{0}.{1}", AssemblyName, Name)
        Dim Stream As System.IO.Stream = Assembly.GetEntryAssembly.GetManifestResourceStream(SreamName)
        Dim Strm As New StreamReader(Stream)
        Return Strm
    End Function
    ''' <summary>
    ''' Using for C#
    ''' </summary>
    ''' <param name="Path"></param>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Shared Function GetStreamReader(Path As String, Name As String) As StreamReader
        If Exists(Name) = False Then Throw New Exception("Embedded Resource Item, Not Found!")
        Dim SPL() As String = Path.Split("/"c, "\"c)
        Dim Location As String = Join(SPL, ".")
        Dim SreamName As String = String.Format("{0}.{1}.{2}", AssemblyName, Location, Name)
        Dim Stream As System.IO.Stream = Assembly.GetEntryAssembly.GetManifestResourceStream(SreamName)
        Dim Strm As New StreamReader(Stream)
        Return Strm
    End Function
    Shared Function GetBytes(Name As String) As Byte()
        If Exists(Name) = False Then Throw New Exception("Embedded Resource Item, Not Found!")
        Dim SreamName As String = String.Format("{0}.{1}", AssemblyName, Name)
        Dim Strm As Stream = Assembly.GetEntryAssembly.GetManifestResourceStream(SreamName)
        Dim buffer(Strm.Length - 1) As Byte
        Strm.Read(buffer, 0, Strm.Length)
        Strm.Close()
        Return buffer
    End Function
    Shared Function GetStream(Name As String) As Stream
        If Exists(Name) = False Then Throw New Exception("Embedded Resource Item, Not Found!")
        Dim SreamName As String = String.Format("{0}.{1}", AssemblyName, Name)
        Dim Strm As Stream = Assembly.GetEntryAssembly.GetManifestResourceStream(SreamName)
        Return Strm
    End Function
    ''' <summary>
    ''' Get Text Document String from Embedded Resource File
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Document(Name As String) As String
        Dim Reader As StreamReader = GetStreamReader(Name)
        Dim Expr As String = Reader.ReadToEnd
        Reader.Close()
        Return Expr
    End Function
    ''' <summary>
    ''' Get Text Document String from Embedded Resource File
    ''' for C#
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Document(Path As String, Name As String) As String
        Dim Reader As StreamReader = GetStreamReader(Path, Name)
        Dim Expr As String = Reader.ReadToEnd
        Reader.Close()
        Return Expr
    End Function
    Shared Browser As New WebBrowser
    ''' <summary>
    ''' Get HTML Document Body InnerHTML from Embedded Resource File
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function HTMLBody(Name As String) As String
        Dim Reader As StreamReader = GetStreamReader(Name)
        Dim Expr As String = Reader.ReadToEnd
        Browser.DocumentText = Expr
        Do Until Browser.ReadyState = WebBrowserReadyState.Complete
            Application.DoEvents()
        Loop
        Reader.Close()
        Return Browser.Document.Body.InnerHtml
    End Function
    Shared Function Iocn(Name As String) As Icon
        Dim Reader As Stream = GetStream(Name)
        Dim Ico As New Icon(Reader)
        Reader.Close()
        Return Ico
    End Function
    Shared Function Bitmap(Name As String) As Bitmap
        Dim Reader As Stream = GetStream(Name)
        Dim Bmp As New Bitmap(Reader)
        Reader.Close()
        Return Bmp
    End Function
    ''' <summary>
    ''' for HTML Image
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Base64StringJpegData(Name As String) As String
        Dim expr As String = String.Format("data:{0};base64,{1}", "image/jpeg", Convert.ToBase64String(GetBytes(Name)))
        Return expr
    End Function
    ''' <summary>
    ''' for HTML Image
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Base64StringPNGData(Name As String) As String
        Dim expr As String = String.Format("data:{0};base64,{1}", "image/png", Convert.ToBase64String(GetBytes(Name)))
        Return expr
    End Function

    ''' <summary>
    ''' for HTML Image
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Base64StringWebpData(Name As String) As String
        Dim expr As String = String.Format("data:{0};base64,{1}", "image/webp", Convert.ToBase64String(GetBytes(Name)))
        Return expr
    End Function
    Shared Function Cursor(Name As String) As Cursor
        Dim Reader As Stream = GetStream(Name)
        Dim cur As New Cursor(Reader)
        Reader.Close()
        Return cur
    End Function

End Class
