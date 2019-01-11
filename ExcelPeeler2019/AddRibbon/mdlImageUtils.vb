Imports System.Windows.Media.Imaging
Imports System.IO



Module MdlImageUtils

    ''' <summary>
    ''' リソースに読み込んだPNGをイメージソースに変換する
    ''' </summary>
    ''' <param name="embeddedPath"></param>
    ''' <returns></returns>
    Public Function BmpImageSource(ByVal embeddedPath As String) As System.Windows.Media.ImageSource

        Dim myAssembly As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim stream As Stream = myAssembly.GetManifestResourceStream(embeddedPath)
        Dim decoder As PngBitmapDecoder = New System.Windows.Media.Imaging.PngBitmapDecoder(stream, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default)

        Return decoder.Frames(0)
    End Function

End Module
