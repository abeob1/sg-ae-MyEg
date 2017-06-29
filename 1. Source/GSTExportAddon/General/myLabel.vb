Imports System.Windows.Forms

Public Class myLabel

    Public Property Caption() As String
        Get
            Return Label1.Text

        End Get
        Set(ByVal value As String)
            Label1.Text = value
        End Set
    End Property


    Public Sub New()

        ' This call is required by the Windows Form Designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.TabStop = False
        'set font


        ' Dim FS As New Font(ThongTin.FontName, ThongTin.FontSise, FontStyle.Regular)
        'Label1.Font = FS

    End Sub

   

End Class
