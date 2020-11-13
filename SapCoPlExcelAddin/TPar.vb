Public Class TPar

    Public aTParCol As Collection
    Public Sub New()
        aTParCol = New Collection
    End Sub

    Public Sub addPar(pParam As String, pVALUE As String, Optional pSTRUCNAME As String = "")
        Dim aStrRec As TStrRec
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        Dim parArray() As String

        If pSTRUCNAME <> "" Then
            aSTRUCNAME = pSTRUCNAME
            aFIELDNAME = pVALUE
        ElseIf InStr(pParam, "-") <> 0 Then
            parArray = Split(pParam, "-")
            aSTRUCNAME = parArray(0)
            aFIELDNAME = parArray(1)
        ElseIf InStr(pParam, "__") <> 0 Then
            parArray = Split(pParam, "_")
            aSTRUCNAME = parArray(0)
            aFIELDNAME = parArray(1)
        End If

        aKey = aSTRUCNAME & "-" & aFIELDNAME
        If aTParCol.Contains(aKey) Then
            aStrRec = aTParCol(aKey)
            aStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE)
        Else
            aStrRec = New TStrRec
            aStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE)
            aTParCol.Add(aStrRec, aKey)
        End If
    End Sub

    Public Function value(pSTRUCNAME As String, pFIELDNAME As String) As String
        Dim aStrRec As TStrRec
        Dim aKey As String
        aKey = pSTRUCNAME & "-" & pFIELDNAME
        If aTParCol.Contains(pSTRUCNAME & "-" & pFIELDNAME) Then
            aStrRec = aTParCol(aKey)
            value = aStrRec.VALUE.Value
        Else
            value = ""
        End If
    End Function

End Class
