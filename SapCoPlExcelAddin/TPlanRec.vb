Public Class TPlanRec

    Public aTPlanRecCol As Collection
    Public Sub New()
        aTPlanRecCol = New Collection
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As TStrRec
        Dim aNameArray() As String
        Dim aKey As String
        Dim aSTRUCNAME As String = ""
        Dim aFIELDNAME As String = ""
        ' do not add empty values
        If Not pEmty And pVALUE = pEmptyChar Then
            Exit Sub
        End If
        If InStr(pNAME, "-") <> 0 Then
            aNameArray = Split(pNAME, "-")
            aSTRUCNAME = aNameArray(0)
            aFIELDNAME = aNameArray(1)
        Else
            aSTRUCNAME = ""
            aFIELDNAME = pNAME
        End If
        aKey = pNAME
        If aTPlanRecCol.Contains(aKey) Then
            aTStrRec = aTPlanRecCol(aKey)
            Select Case pOperation
                Case "add"
                    aTStrRec.addValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "sub"
                    aTStrRec.subValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "mul"
                    aTStrRec.mulValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case "div"
                    aTStrRec.divValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
                Case Else
                    aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            End Select
        Else
            aTStrRec = New TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            aTPlanRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTPlanRec As TPlanRec)
        Dim aTStrRec As TStrRec
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            setValues(aTStrRec.getKey(), aTStrRec.VALUE.Value, aTStrRec.CURRENCY.Value, aTStrRec.FORMAT.Value)
        Next
    End Sub

    Public Sub addValues(pTPlanRec As TPlanRec)
        Dim aTStrRec As TStrRec
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            If aTStrRec.CURRENCY.Value <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.VALUE.Value, aTStrRec.CURRENCY.Value, aTStrRec.FORMAT.Value, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.VALUE.Value, aTStrRec.CURRENCY.Value, aTStrRec.FORMAT.Value)
            End If
        Next
    End Sub

    Public Sub buildValFields()
        Dim aPAIQv As String = ""
        Dim aPAIQf As String = ""
        Dim aPPCQv As String = ""
        Dim aPPCQf As String = ""
        Dim aPPCVv As String = ""
        Dim aPPCVf As String = ""
        ' get the period field for all values
        Dim aPerAll As TStrRec = Nothing
        If aTPlanRecCol.Contains("TMPPERVALUE-ALL_PERXX") Then
            aPerAll = aTPlanRecCol("TMPPERVALUE-ALL_PERXX")
            aPAIQv = Format(CInt(aPerAll.VALUE.Value), "00")
            aPAIQf = aPAIQv
            aPPCQv = aPAIQv
            aPPCQf = aPAIQv
            aPPCVv = aPAIQv
            aPPCVf = aPAIQv
        End If
        ' Only get the other period fields if ALL_PAR is nothing
        If aPerAll Is Nothing Then
            ' AI-Quantities
            Dim aPerAIQv As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_VAR_PERXX") Then
                aPerAIQv = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQv = Format(CInt(aPerAIQv.VALUE.Value), "00")
            End If
            Dim aPerAIQf As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_FIX_PERXX") Then
                aPerAIQf = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQf = Format(CInt(aPerAIQf.VALUE.Value), "00")
            End If
            ' PC-Quantities
            Dim aPerPCQv As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_QUAN_PERXX") Then
                aPerPCQv = aTPlanRecCol("TMPPERVALUE-VAR_QUAN_PERXX")
                aPPCQv = Format(CInt(aPerPCQv.VALUE.Value), "00")
            End If
            Dim aPerPCQf As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_QUAN_PERXX") Then
                aPerPCQf = aTPlanRecCol("TMPPERVALUE-FIX_QUAN_PERXX")
                aPPCQf = Format(CInt(aPerPCQf.VALUE.Value), "00")
            End If
            ' PC-Values
            Dim aPerPCVv As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_VAL_PERXX") Then
                aPerPCVv = aTPlanRecCol("TMPPERVALUE-VAR_VAL_PERXX")
                aPPCVv = Format(CInt(aPerPCVv.VALUE.Value), "00")
            End If
            Dim aPerPCVf As TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_VAL_PERXX") Then
                aPerPCVf = aTPlanRecCol("TMPPERVALUE-FIX_VAL_PERXX")
                aPPCVf = Format(CInt(aPerPCVf.VALUE.Value), "00")
            End If
        End If
        If aPAIQv <> "" Then
            addNewValueField("TMPPERVALUE-QUANTITY_VAR", aPAIQv)
        End If
        If aPAIQf <> "" Then
            addNewValueField("TMPPERVALUE-QUANTITY_FIX", aPAIQf)
        End If
        If aPPCQv <> "" Then
            addNewValueField("TMPPERVALUE-VAR_QUAN", aPPCQv)
        End If
        If aPPCQf <> "" Then
            addNewValueField("TMPPERVALUE-FIX_QUAN", aPPCQf)
        End If
        If aPPCVv <> "" Then
            addNewValueField("TMPPERVALUE-VAR_VAL", aPPCVv)
        End If
        If aPPCVf <> "" Then
            addNewValueField("TMPPERVALUE-FIX_VAL", aPPCVf)
        End If
    End Sub

    Private Sub addNewValueField(aTmpName As String, aPer As String)
        Dim aTmpPlanVal As TStrRec
        Dim aNewName As String
        If aTPlanRecCol.Contains(aTmpName) Then
            aTmpPlanVal = aTPlanRecCol(aTmpName)
            aNewName = Right(aTmpPlanVal.STRUCNAME.Value, Len(aTmpPlanVal.STRUCNAME.Value) - 3) & "-" & aTmpPlanVal.FIELDNAME.Value & "_PER" & aPer
            setValues(aNewName, aTmpPlanVal.VALUE.Value, aTmpPlanVal.CURRENCY.Value, aTmpPlanVal.FORMAT.Value)
        End If
    End Sub

    Public Sub checkWBSLimit(ByRef pWBSLimit As WBSLimit, ByRef pPar As TPar)
        If Not pWBSLimit.checkWBSLimit(getWBS(pPar), getYear(pPar), getPer(pPar)) Then
            setAmount("0", pPar)
            setQuantity("0", pPar)
        End If
    End Sub

    Public Sub mapCE(ByRef pCeMap As CeMap, ByRef pPar As TPar)
        Dim aCE As String
        Dim aNewCE As String
        aCE = getCE(pPar)
        aNewCE = pCeMap.map(aCE)
        If aCE <> aNewCE Then
            setCE(aNewCE, pPar)
        End If
    End Sub

    Public Function getYear(ByRef pPar As TPar) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTFISCYEAR") <> "", pPar.value("COL", "PSIMPORTFISCYEAR"), "HEADERINFO-FISC_YEAR")
        Dim aTStrRec As TStrRec
        getYear = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getYear = aTStrRec.VALUE.Value
        End If
    End Function

    Public Function getPer(ByRef pPar As TPar) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTFISCPER") <> "", pPar.value("COL", "PSIMPORTFISCPER"), "TMPPERVALUE-ALL_PERXX")
        Dim aTStrRec As TStrRec
        getPer = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getPer = aTStrRec.VALUE.Value
        End If
    End Function

    Public Function getWBS(ByRef pPar As TPar) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTWBS") <> "", pPar.value("COL", "PSIMPORTWBS"), "COOBJECT-WBS_ELEMENT")
        Dim aTStrRec As TStrRec
        getWBS = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getWBS = aTStrRec.VALUE.Value
        End If
    End Function

    Public Function getCE(ByRef pPar As TPar) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTCE") <> "", pPar.value("COL", "PSIMPORTCE"), "PERVALUE-COST_ELEM")
        Dim aTStrRec As TStrRec
        getCE = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getCE = aTStrRec.VALUE.Value
        End If
    End Function
    Public Sub setCE(pValue As String, ByRef pPar As TPar)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTCE") <> "", pPar.value("COL", "PSIMPORTCE"), "PERVALUE-COST_ELEM")
        Dim aTStrRec As TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.VALUE.Value = pValue
        End If
    End Sub

    Public Sub setAmount(pValue As String, ByRef pPar As TPar)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTAMOUNT") <> "", pPar.value("COL", "PSIMPORTAMOUNT"), "TMPPERVALUE-VAR_VAL")
        Dim aTStrRec As TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.VALUE.Value = pValue
        End If
    End Sub

    Public Sub setQuantity(pValue As String, ByRef pPar As TPar)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTQUANTITY") <> "", pPar.value("COL", "PSIMPORTQUANTITY"), "TMPPERVALUE-QUANTITY_VAR")
        Dim aTStrRec As TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.VALUE.Value = pValue
        End If
    End Sub

End Class