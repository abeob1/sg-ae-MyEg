Module PublicVariable
    'Connection
    Public oCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company

    Public WithEvents SBO_Application As SAPbouiCOM.Application

    Public SAPPass As String = ""

    Public Version As String = "20161111"

    Public IsDebug As String = "N"
End Module
