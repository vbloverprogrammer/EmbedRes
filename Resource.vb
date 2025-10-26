Imports System.Reflection
Public Class Resource
    Public Shared ReadOnly Property AssemblyName() As String
        Get
            Return Assembly.GetEntryAssembly.GetName.Name.Replace(Space(1), "_")
        End Get
    End Property
    Private Shared ReadOnly Property ResourceManager(ResourcesName As String) As Global.System.Resources.ResourceManager
        Get
            Dim ResourceName As String = String.Format("{0}.{1}", AssemblyName, ResourcesName)
            Return New Global.System.Resources.ResourceManager(ResourceName, Assembly.GetEntryAssembly)
        End Get
    End Property
    ''' <summary>
    ''' for HTML Image
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Base64StringJpegData(Resources As String, Name As String) As String
        Dim Bitmap As System.Drawing.Bitmap = ResourceManager(Resources).GetObject(Name)
        Dim expr As String
        Using MStream As New IO.MemoryStream()
            Bitmap.Save(MStream, System.Drawing.Imaging.ImageFormat.Jpeg)
            expr = String.Format("data:{0};base64,{1}", "image/jpeg", Convert.ToBase64String(MStream.ToArray))
            Bitmap.Clone()
            MStream.Close()
        End Using
        Return expr
    End Function
    ''' <summary>
    ''' for HTML Image
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <returns></returns>
    Shared Function Base64StringPNGData(Resources As String, Name As String) As String
        Dim Bitmap As System.Drawing.Bitmap = ResourceManager(Resources).GetObject(Name)
        Dim expr As String
        Using MStream As New IO.MemoryStream()
            Bitmap.Save(MStream, System.Drawing.Imaging.ImageFormat.Png)
            expr = String.Format("data:{0};base64,{1}", "image/png", Convert.ToBase64String(MStream.ToArray))
            Bitmap.Clone()
            MStream.Close()
        End Using
        Return expr
    End Function
End Class
