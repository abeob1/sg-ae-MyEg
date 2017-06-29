Imports System.Collections.Generic
Imports System.Text
Imports SAPbouiCOM

Public Class clDrawItem
    Class ItemLayout
        Public Sub New(ByVal iLeft As Integer, ByVal iTop As Integer, ByVal iHeight As Integer, ByVal iWidth As Integer)
            Me.m_left = iLeft
            Me.m_top = iTop
            Me.m_height = iHeight
            Me.m_width = iWidth
        End Sub

        Public Sub New(ByVal layout As ItemLayout)
            Me.m_left = layout.Left
            Me.m_top = layout.Top
            Me.m_height = layout.Height
            Me.m_width = layout.Width
        End Sub

        'Left property
        Private m_left As Integer

        Public Property Left() As Integer
            Get
                Return m_left
            End Get
            Set(ByVal value As Integer)
                m_left = value
            End Set
        End Property

        'Top property
        Private m_top As Integer

        Public Property Top() As Integer
            Get
                Return m_top
            End Get
            Set(ByVal value As Integer)
                m_top = value
            End Set
        End Property

        'Height property
        Private m_height As Integer

        Public Property Height() As Integer
            Get
                Return m_height
            End Get
            Set(ByVal value As Integer)
                m_height = value
            End Set
        End Property

        'Width property
        Private m_width As Integer

        Public Property Width() As Integer
            Get
                Return m_width
            End Get
            Set(ByVal value As Integer)
                m_width = value
            End Set
        End Property
    End Class

    Class FormItemCreator
        Public Shared Function CreateItem(ByVal oForm As Form, ByVal uid As String, ByVal itemType As SAPbouiCOM.BoFormItemTypes, ByVal layout As ItemLayout) As SAPbouiCOM.Item
            Dim oItem As SAPbouiCOM.Item = oForm.Items.Add(uid, itemType)
            oItem.Top = layout.Top
            oItem.Left = layout.Left
            oItem.Height = layout.Height
            oItem.Width = layout.Width

            Return oItem
        End Function
    End Class
End Class
