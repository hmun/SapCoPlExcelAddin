Imports SAP.Middleware.Connector

Public Class SapRFCGetMeta

    Private Shared ReadOnly log As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private oRfcFunction As IRfcFunction
    Private destination As RfcCustomDestination
    Private sapcon As SapCon
    Private par As SAPCommon.TStr
    Private cName As String = "SAPCostActivityPlanning"

    Sub New(aSapCon As SapCon, ByRef aPar As SAPCommon.TStr)
        Try
            sapcon = aSapCon
            par = aPar
            aSapCon.getDestination(destination)
        Catch ex As System.Exception
            log.Error("New - Exception=" & ex.ToString)
            MsgBox("New failed! " & ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        End Try
    End Sub

    Private Sub addToStrucDic(pArrayName As String, pRfcStructureMetadata As RfcStructureMetadata, ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        If pStrucDic.ContainsKey(pArrayName) Then
            pStrucDic.Remove(pArrayName)
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pStrucDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Private Sub addToFieldDic(pArrayName As String, pRfcStructureMetadata As RfcParameterMetadata, ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata))
        If pFieldDic.ContainsKey(pArrayName) Then
            pFieldDic.Remove(pArrayName)
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        Else
            pFieldDic.Add(pArrayName, pRfcStructureMetadata)
        End If
    End Sub

    Public Sub getMeta(pFunction As String, pStructures As String(), pImports As String(), pTables As String(),
                                    ByRef pFieldDic As Dictionary(Of String, RfcParameterMetadata), ByRef pStrucDic As Dictionary(Of String, RfcStructureMetadata))
        Try
            log.Debug("getMeta - " & "creating Function " & pFunction)
            oRfcFunction = destination.Repository.CreateFunction(pFunction)
            Dim oStructure As IRfcStructure
            Dim oTable As IRfcTable
            ' Imports
            For s As Integer = 0 To pImports.Length - 1
                Try
                    addToFieldDic("I|" & pImports(s), oRfcFunction.Metadata.Item(pImports(s)), pFieldDic)
                Catch Ex As System.Exception
                    log.Debug("getMeta_PostActivityOutput - Object does not exist: " & pImports(s))
                End Try
            Next
            ' Import Strcutures
            For s As Integer = 0 To pStructures.Length - 1
                Try
                    oStructure = oRfcFunction.GetStructure(pStructures(s))
                    addToStrucDic("S|" & pStructures(s), oStructure.Metadata, pStrucDic)
                Catch Ex As System.Exception
                    log.Debug("getMeta_PostActivityOutput - Object does not exist: " & pStructures(s))
                End Try
            Next
            For s As Integer = 0 To pTables.Length - 1
                Try
                    oTable = oRfcFunction.GetTable(pTables(s))
                    addToStrucDic("T|" & pTables(s), oTable.Metadata.LineType, pStrucDic)
                Catch Ex As System.Exception
                    log.Debug("getMeta_PostActivityOutput - Object does not exist: " & pTables(s))
                End Try
            Next
        Catch Ex As System.Exception
            log.Error("getMeta - Exception=" & Ex.ToString)
            MsgBox("Error: Exception " & Ex.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Critical, cName)
        Finally
            log.Debug("getMeta - " & "EndContext")
            RfcSessionManager.EndContext(destination)
        End Try
    End Sub

End Class
