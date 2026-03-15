Public Class TPlanRec

    Public aTPlanRecCol As Collection
    Public Sub New()
        aTPlanRecCol = New Collection
    End Sub

    Public Sub setValues(pNAME As String, pVALUE As String, pCURRENCY As String, pFORMAT As String,
                         Optional pEmty As Boolean = False, Optional pEmptyChar As String = "#", Optional pOperation As String = "set")
        Dim aTStrRec As SAPCommon.TStrRec
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
            aTStrRec = New SAPCommon.TStrRec
            aTStrRec.setValues(aSTRUCNAME, aFIELDNAME, pVALUE, pCURRENCY, pFORMAT)
            aTPlanRecCol.Add(aTStrRec, aKey)
        End If
    End Sub

    Public Sub setValues(pTPlanRec As TPlanRec)
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
        Next
    End Sub

    Public Sub addValues(pTPlanRec As TPlanRec)
        Dim aTStrRec As SAPCommon.TStrRec
        For Each aTStrRec In pTPlanRec.aTPlanRecCol
            If aTStrRec.Currency <> "" Then
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format, pOperation:="add")
            Else
                setValues(aTStrRec.getKey(), aTStrRec.Value, aTStrRec.Currency, aTStrRec.Format)
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
        Dim aPerAll As SAPCommon.TStrRec = Nothing
        If aTPlanRecCol.Contains("TMPPERVALUE-ALL_PERXX") Then
            aPerAll = aTPlanRecCol("TMPPERVALUE-ALL_PERXX")
            aPAIQv = Format(CInt(aPerAll.Value), "00")
            aPAIQf = aPAIQv
            aPPCQv = aPAIQv
            aPPCQf = aPAIQv
            aPPCVv = aPAIQv
            aPPCVf = aPAIQv
        End If
        ' Only get the other period fields if ALL_PAR is nothing
        If aPerAll Is Nothing Then
            ' AI-Quantities
            Dim aPerAIQv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_VAR_PERXX") Then
                aPerAIQv = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQv = Format(CInt(aPerAIQv.Value), "00")
            End If
            Dim aPerAIQf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_FIX_PERXX") Then
                aPerAIQf = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQf = Format(CInt(aPerAIQf.Value), "00")
            End If
            ' PC-Quantities
            Dim aPerPCQv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_QUAN_PERXX") Then
                aPerPCQv = aTPlanRecCol("TMPPERVALUE-VAR_QUAN_PERXX")
                aPPCQv = Format(CInt(aPerPCQv.Value), "00")
            End If
            Dim aPerPCQf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_QUAN_PERXX") Then
                aPerPCQf = aTPlanRecCol("TMPPERVALUE-FIX_QUAN_PERXX")
                aPPCQf = Format(CInt(aPerPCQf.Value), "00")
            End If
            ' PC-Values
            Dim aPerPCVv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_VAL_PERXX") Then
                aPerPCVv = aTPlanRecCol("TMPPERVALUE-VAR_VAL_PERXX")
                aPPCVv = Format(CInt(aPerPCVv.Value), "00")
            End If
            Dim aPerPCVf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_VAL_PERXX") Then
                aPerPCVf = aTPlanRecCol("TMPPERVALUE-FIX_VAL_PERXX")
                aPPCVf = Format(CInt(aPerPCVf.Value), "00")
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

    Public Sub buildValFieldsTot(pPeriod As String)
        Dim aPAIQv As String = ""
        Dim aPAIQf As String = ""
        Dim aPPCQv As String = ""
        Dim aPPCQf As String = ""
        Dim aPPCVv As String = ""
        Dim aPPCVf As String = ""
        ' get the period field for all values
        Dim aPerAll As SAPCommon.TStrRec = Nothing
        If aTPlanRecCol.Contains("TMPPERVALUE-ALL_PERXX") Then
            aPerAll = aTPlanRecCol("TMPPERVALUE-ALL_PERXX")
            aPAIQv = pPeriod
            aPAIQf = aPAIQv
            aPPCQv = aPAIQv
            aPPCQf = aPAIQv
            aPPCVv = aPAIQv
            aPPCVf = aPAIQv
        End If
        ' Only get the other period fields if ALL_PAR is nothing
        If aPerAll Is Nothing Then
            ' AI-Quantities
            Dim aPerAIQv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_VAR_PERXX") Then
                aPerAIQv = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQv = pPeriod
            End If
            Dim aPerAIQf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-QUANTITY_FIX_PERXX") Then
                aPerAIQf = aTPlanRecCol("TMPPERVALUE-QUANTITY_VAR_PERXX")
                aPAIQf = pPeriod
            End If
            ' PC-Quantities
            Dim aPerPCQv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_QUAN_PERXX") Then
                aPerPCQv = aTPlanRecCol("TMPPERVALUE-VAR_QUAN_PERXX")
                aPPCQv = pPeriod
            End If
            Dim aPerPCQf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_QUAN_PERXX") Then
                aPerPCQf = aTPlanRecCol("TMPPERVALUE-FIX_QUAN_PERXX")
                aPPCQf = pPeriod
            End If
            ' PC-Values
            Dim aPerPCVv As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-VAR_VAL_PERXX") Then
                aPerPCVv = aTPlanRecCol("TMPPERVALUE-VAR_VAL_PERXX")
                aPPCVv = pPeriod
            End If
            Dim aPerPCVf As SAPCommon.TStrRec = Nothing
            If aTPlanRecCol.Contains("TMPPERVALUE-FIX_VAL_PERXX") Then
                aPerPCVf = aTPlanRecCol("TMPPERVALUE-FIX_VAL_PERXX")
                aPPCVf = pPeriod
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
        Dim aTmpPlanVal As SAPCommon.TStrRec
        Dim aNewName As String
        If aTPlanRecCol.Contains(aTmpName) Then
            aTmpPlanVal = aTPlanRecCol(aTmpName)
            aNewName = Right(aTmpPlanVal.Strucname, Len(aTmpPlanVal.Strucname) - 3) & "-" & aTmpPlanVal.Fieldname & "_PER" & aPer
            setValues(aNewName, aTmpPlanVal.Value, aTmpPlanVal.Currency, aTmpPlanVal.Format)
        End If
    End Sub

    Public Sub checkWBSLimit(ByRef pWBSLimit As WBSLimit, ByRef pPar As SAPCommon.TStr)
        If Not pWBSLimit.checkWBSLimit(getWBS(pPar), getYear(pPar), getPer(pPar)) Then
            setAmount("0", pPar)
            setQuantity("0", pPar)
        End If
    End Sub

    Public Sub mapCE(ByRef pCeMap As CeMap, ByRef pPar As SAPCommon.TStr)
        Dim aCE As String
        Dim aNewCE As String
        aCE = getCE(pPar)
        aNewCE = pCeMap.map(aCE)
        If aCE <> aNewCE Then
            setCE(aNewCE, pPar)
        End If
    End Sub

    Public Function getYear(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTFISCYEAR") <> "", pPar.value("COL", "PSIMPORTFISCYEAR"), "HEADERINFO-FISC_YEAR")
        Dim aTStrRec As SAPCommon.TStrRec
        getYear = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getYear = aTStrRec.Value
        End If
    End Function

    Public Function getPer(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTFISCPER") <> "", pPar.value("COL", "PSIMPORTFISCPER"), "TMPPERVALUE-ALL_PERXX")
        Dim aTStrRec As SAPCommon.TStrRec
        getPer = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getPer = aTStrRec.Value
        End If
    End Function

    Public Function getWBS(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTWBS") <> "", pPar.value("COL", "PSIMPORTWBS"), "COOBJECT-WBS_ELEMENT")
        Dim aTStrRec As SAPCommon.TStrRec
        getWBS = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getWBS = aTStrRec.Value
        End If
    End Function

    Public Function getCE(ByRef pPar As SAPCommon.TStr) As String
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTCE") <> "", pPar.value("COL", "PSIMPORTCE"), "PERVALUE-COST_ELEM")
        Dim aTStrRec As SAPCommon.TStrRec
        getCE = ""
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            getCE = aTStrRec.Value
        End If
    End Function
    Public Sub setCE(pValue As String, ByRef pPar As SAPCommon.TStr)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTCE") <> "", pPar.value("COL", "PSIMPORTCE"), "PERVALUE-COST_ELEM")
        Dim aTStrRec As SAPCommon.TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.Value = pValue
        End If
    End Sub

    Public Sub setAmount(pValue As String, ByRef pPar As SAPCommon.TStr)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTAMOUNT") <> "", pPar.value("COL", "PSIMPORTAMOUNT"), "TMPPERVALUE-VAR_VAL")
        Dim aTStrRec As SAPCommon.TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.Value = pValue
        End If
    End Sub

    Public Sub setQuantity(pValue As String, ByRef pPar As SAPCommon.TStr)
        Dim aClmn As String = If(pPar.value("COL", "PSIMPORTQUANTITY") <> "", pPar.value("COL", "PSIMPORTQUANTITY"), "TMPPERVALUE-QUANTITY_VAR")
        Dim aTStrRec As SAPCommon.TStrRec
        If aTPlanRecCol.Contains(aClmn) Then
            aTStrRec = aTPlanRecCol(aClmn)
            aTStrRec.Value = pValue
        End If
    End Sub

End Class