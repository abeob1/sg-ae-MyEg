Imports System.Data
Imports System.Xml
Public Class clXml2Ds
    Public Function xmlStr2Ds(xmlString As String, Optional multipleXML As Boolean = False, Optional ObjType As String = "") As DataSet
        Dim dsBO As New DataSet()
        Dim xml As New XmlDocument()
        xml.LoadXml(xmlString)
        BuildDataSet(dsBO, xml.ChildNodes, multipleXML, ObjType)
        Return dsBO
    End Function
    Private Sub BuildDataSet(ds As DataSet, parentNode As XmlNodeList, multipleXML As Boolean, ObjType As String)
        For Each nodeParent As XmlNode In parentNode
            If nodeParent.Name.Equals("BO") Or (ObjType = "25" And nodeParent.Name.Equals("Deposit")) Then
                For Each node As XmlNode In nodeParent.ChildNodes
                    If node.Name.Equals("AdmInfo") Then
                        Continue For
                    End If
                    BuildTable(ds, node)
                Next
                If Not multipleXML Then
                    Exit For
                End If
            Else
                BuildDataSet(ds, nodeParent.ChildNodes, multipleXML, ObjType)
            End If
        Next
    End Sub
    Private Sub BuildTable(ads As DataSet, nodeTable As XmlNode)
        Dim dt As New DataTable(nodeTable.Name)
        Dim firstNode As XmlNode = nodeTable.FirstChild
        If Not IsNothing(firstNode) Then
            For Each col As XmlNode In firstNode.ChildNodes
                If Not dt.Columns.Contains(col.Name) Then
                    dt.Columns.Add(New DataColumn(col.Name, Type.[GetType]("System.String")))
                End If
            Next
        End If

        For Each rowNode As XmlNode In nodeTable.ChildNodes
            Dim dr As DataRow = dt.NewRow()
            For Each colValue As XmlNode In rowNode.ChildNodes
                If Not dt.Columns.Contains(colValue.Name) Then
                    dt.Columns.Add(colValue.Name)
                End If
                dr(colValue.Name) = colValue.InnerText
            Next
            dt.Rows.Add(dr)
        Next
        If Not ads.Tables.Contains(dt.TableName) Then
            ads.Tables.Add(dt.Copy)
        End If

    End Sub
End Class
