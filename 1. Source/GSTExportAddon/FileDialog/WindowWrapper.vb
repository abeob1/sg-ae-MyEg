Imports System
Imports System.Collections.Generic
Imports System.Text


Public Class WindowWrapper
    Implements System.Windows.Forms.IWin32Window
    Private _hwnd As IntPtr

    ' Property
    Public Overridable ReadOnly Property Handle() As IntPtr Implements System.Windows.Forms.IWin32Window.Handle
        Get
            Return _hwnd
        End Get
    End Property

    ' Constructor
    Public Sub New(handle As IntPtr)
        _hwnd = handle
    End Sub
End Class