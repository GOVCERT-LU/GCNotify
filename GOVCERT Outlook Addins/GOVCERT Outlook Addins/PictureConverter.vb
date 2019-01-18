' Source: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/aa338202(v=office.12)'

Imports System.Windows.Forms

'Prevention from opening with the designer'
<System.ComponentModel.DesignerCategory("")>
Public Class PictureConverter
    Inherits AxHost

    Private Sub New()
        MyBase.New(String.Empty)
    End Sub

    Public Shared Function ImageToPictureDisp(ByVal image As System.Drawing.Image) As stdole.IPictureDisp
        Return CType(GetIPictureDispFromPicture(image), stdole.IPictureDisp)
    End Function

    Public Shared Function IconToPictureDisp(ByVal icon As System.Drawing.Icon) As stdole.IPictureDisp
        Return ImageToPictureDisp(icon.ToBitmap())
    End Function

    Public Shared Function PictureDispToImage(ByVal picture As stdole.IPictureDisp) As System.Drawing.Image
        Return GetPictureFromIPicture(picture)
    End Function

End Class
