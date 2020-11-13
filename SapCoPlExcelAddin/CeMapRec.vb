Public Class CeMapRec

    Public CE As SAPCommon.TField
    Public NewCE As SAPCommon.TField

    Public Function setValues(pCE As String, pNewCE As String)
        CE = New SAPCommon.TField("CE", pCE)
        NewCE = New SAPCommon.TField("NewCE", pNewCE)
    End Function

    Public Function getKey() As String
        Dim aKey As String
        aKey = CE.Value
        getKey = aKey
    End Function

    Public Function getKeyR() As String
        Dim aKey As String
        aKey = CE.Value
        getKeyR = aKey
    End Function

    Public Function toStringValue() As Object
        Dim aArray(1) As String
        aArray(0) = CE.Value
        aArray(1) = NewCE.Value
        toStringValue = aArray
    End Function

End Class
