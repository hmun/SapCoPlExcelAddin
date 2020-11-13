Partial Class SapCoPlRibbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapCoPlRibbon))
        Me.SapCoPl = Me.Factory.CreateRibbonTab
        Me.SAPCoOmPlan = Me.Factory.CreateRibbonGroup
        Me.ButtonReadAO = Me.Factory.CreateRibbonButton
        Me.ButtonReadPC = Me.Factory.CreateRibbonButton
        Me.ButtonReadAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostAI = Me.Factory.CreateRibbonButton
        Me.ButtonReadSK = Me.Factory.CreateRibbonButton
        Me.ButtonPostSK = Me.Factory.CreateRibbonButton
        Me.SAPCoOmPlanPer = Me.Factory.CreateRibbonGroup
        Me.ButtonReadPerAO = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerPC = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerAI = Me.Factory.CreateRibbonButton
        Me.ButtonReadPerSK = Me.Factory.CreateRibbonButton
        Me.ButtonPostPerSK = Me.Factory.CreateRibbonButton
        Me.SAPPsPlan = Me.Factory.CreateRibbonGroup
        Me.ButtonPsUpdCheck = Me.Factory.CreateRibbonButton
        Me.ButtonPsUpdPost = Me.Factory.CreateRibbonButton
        Me.SAPCoCosting = Me.Factory.CreateRibbonGroup
        Me.ButtonCostingCreate = Me.Factory.CreateRibbonButton
        Me.ButtonCostingChange = Me.Factory.CreateRibbonButton
        Me.SAPCoPlLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapCoPl.SuspendLayout()
        Me.SAPCoOmPlan.SuspendLayout()
        Me.SAPCoOmPlanPer.SuspendLayout()
        Me.SAPPsPlan.SuspendLayout()
        Me.SAPCoCosting.SuspendLayout()
        Me.SAPCoPlLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCoPl
        '
        Me.SapCoPl.Groups.Add(Me.SAPCoOmPlan)
        Me.SapCoPl.Groups.Add(Me.SAPCoOmPlanPer)
        Me.SapCoPl.Groups.Add(Me.SAPPsPlan)
        Me.SapCoPl.Groups.Add(Me.SAPCoCosting)
        Me.SapCoPl.Groups.Add(Me.SAPCoPlLogon)
        Me.SapCoPl.Label = "SAP CO-PL"
        Me.SapCoPl.Name = "SapCoPl"
        Me.SapCoPl.Position = Me.Factory.RibbonPosition.AfterOfficeId("SapCo")
        '
        'SAPCoOmPlan
        '
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadAO)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadPC)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadAI)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostAO)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostPC)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostAI)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonReadSK)
        Me.SAPCoOmPlan.Items.Add(Me.ButtonPostSK)
        Me.SAPCoOmPlan.Label = "CO-OM Plan"
        Me.SAPCoOmPlan.Name = "SAPCoOmPlan"
        '
        'ButtonReadAO
        '
        Me.ButtonReadAO.Image = CType(resources.GetObject("ButtonReadAO.Image"), System.Drawing.Image)
        Me.ButtonReadAO.Label = "Read AO"
        Me.ButtonReadAO.Name = "ButtonReadAO"
        Me.ButtonReadAO.ScreenTip = "Read Activity Output"
        Me.ButtonReadAO.ShowImage = True
        '
        'ButtonReadPC
        '
        Me.ButtonReadPC.Image = CType(resources.GetObject("ButtonReadPC.Image"), System.Drawing.Image)
        Me.ButtonReadPC.Label = "Read PC"
        Me.ButtonReadPC.Name = "ButtonReadPC"
        Me.ButtonReadPC.ScreenTip = "Read Primary Cost"
        Me.ButtonReadPC.ShowImage = True
        '
        'ButtonReadAI
        '
        Me.ButtonReadAI.Image = CType(resources.GetObject("ButtonReadAI.Image"), System.Drawing.Image)
        Me.ButtonReadAI.Label = "Read AI"
        Me.ButtonReadAI.Name = "ButtonReadAI"
        Me.ButtonReadAI.ScreenTip = "Read Activity Input"
        Me.ButtonReadAI.ShowImage = True
        '
        'ButtonPostAO
        '
        Me.ButtonPostAO.Image = CType(resources.GetObject("ButtonPostAO.Image"), System.Drawing.Image)
        Me.ButtonPostAO.Label = "Post AO"
        Me.ButtonPostAO.Name = "ButtonPostAO"
        Me.ButtonPostAO.ScreenTip = "Post Activity Output"
        Me.ButtonPostAO.ShowImage = True
        '
        'ButtonPostPC
        '
        Me.ButtonPostPC.Image = CType(resources.GetObject("ButtonPostPC.Image"), System.Drawing.Image)
        Me.ButtonPostPC.Label = "Post PC"
        Me.ButtonPostPC.Name = "ButtonPostPC"
        Me.ButtonPostPC.ScreenTip = "Post Primary Cost"
        Me.ButtonPostPC.ShowImage = True
        '
        'ButtonPostAI
        '
        Me.ButtonPostAI.Image = CType(resources.GetObject("ButtonPostAI.Image"), System.Drawing.Image)
        Me.ButtonPostAI.Label = "Post AI"
        Me.ButtonPostAI.Name = "ButtonPostAI"
        Me.ButtonPostAI.ScreenTip = "Post Activity Input"
        Me.ButtonPostAI.ShowImage = True
        '
        'ButtonReadSK
        '
        Me.ButtonReadSK.Image = CType(resources.GetObject("ButtonReadSK.Image"), System.Drawing.Image)
        Me.ButtonReadSK.Label = "Read SK"
        Me.ButtonReadSK.Name = "ButtonReadSK"
        Me.ButtonReadSK.ScreenTip = "Read Statistical Keyfigures"
        Me.ButtonReadSK.ShowImage = True
        '
        'ButtonPostSK
        '
        Me.ButtonPostSK.Image = CType(resources.GetObject("ButtonPostSK.Image"), System.Drawing.Image)
        Me.ButtonPostSK.Label = "Post SK"
        Me.ButtonPostSK.Name = "ButtonPostSK"
        Me.ButtonPostSK.ScreenTip = "Post Statistical Keyfigures"
        Me.ButtonPostSK.ShowImage = True
        '
        'SAPCoOmPlanPer
        '
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerAO)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerPC)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerAI)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerAO)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerPC)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerAI)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonReadPerSK)
        Me.SAPCoOmPlanPer.Items.Add(Me.ButtonPostPerSK)
        Me.SAPCoOmPlanPer.Label = "CO-OM Plan Periodic"
        Me.SAPCoOmPlanPer.Name = "SAPCoOmPlanPer"
        '
        'ButtonReadPerAO
        '
        Me.ButtonReadPerAO.Image = CType(resources.GetObject("ButtonReadPerAO.Image"), System.Drawing.Image)
        Me.ButtonReadPerAO.Label = "Read Per AO"
        Me.ButtonReadPerAO.Name = "ButtonReadPerAO"
        Me.ButtonReadPerAO.ScreenTip = "Read Activity Output"
        Me.ButtonReadPerAO.ShowImage = True
        '
        'ButtonReadPerPC
        '
        Me.ButtonReadPerPC.Image = CType(resources.GetObject("ButtonReadPerPC.Image"), System.Drawing.Image)
        Me.ButtonReadPerPC.Label = "Read Per PC"
        Me.ButtonReadPerPC.Name = "ButtonReadPerPC"
        Me.ButtonReadPerPC.ScreenTip = "Read Primary Cost"
        Me.ButtonReadPerPC.ShowImage = True
        '
        'ButtonReadPerAI
        '
        Me.ButtonReadPerAI.Image = CType(resources.GetObject("ButtonReadPerAI.Image"), System.Drawing.Image)
        Me.ButtonReadPerAI.Label = "Read Per AI"
        Me.ButtonReadPerAI.Name = "ButtonReadPerAI"
        Me.ButtonReadPerAI.ScreenTip = "Read Activity Input"
        Me.ButtonReadPerAI.ShowImage = True
        '
        'ButtonPostPerAO
        '
        Me.ButtonPostPerAO.Image = CType(resources.GetObject("ButtonPostPerAO.Image"), System.Drawing.Image)
        Me.ButtonPostPerAO.Label = "Post Per AO"
        Me.ButtonPostPerAO.Name = "ButtonPostPerAO"
        Me.ButtonPostPerAO.ScreenTip = "Post Activity Output"
        Me.ButtonPostPerAO.ShowImage = True
        '
        'ButtonPostPerPC
        '
        Me.ButtonPostPerPC.Image = CType(resources.GetObject("ButtonPostPerPC.Image"), System.Drawing.Image)
        Me.ButtonPostPerPC.Label = "Post Per PC"
        Me.ButtonPostPerPC.Name = "ButtonPostPerPC"
        Me.ButtonPostPerPC.ScreenTip = "Post Primary Cost"
        Me.ButtonPostPerPC.ShowImage = True
        '
        'ButtonPostPerAI
        '
        Me.ButtonPostPerAI.Image = CType(resources.GetObject("ButtonPostPerAI.Image"), System.Drawing.Image)
        Me.ButtonPostPerAI.Label = "Post Per AI"
        Me.ButtonPostPerAI.Name = "ButtonPostPerAI"
        Me.ButtonPostPerAI.ScreenTip = "Post Activity Input"
        Me.ButtonPostPerAI.ShowImage = True
        '
        'ButtonReadPerSK
        '
        Me.ButtonReadPerSK.Image = CType(resources.GetObject("ButtonReadPerSK.Image"), System.Drawing.Image)
        Me.ButtonReadPerSK.Label = "Read Per SK"
        Me.ButtonReadPerSK.Name = "ButtonReadPerSK"
        Me.ButtonReadPerSK.ScreenTip = "Read Statistical Keyfigures"
        Me.ButtonReadPerSK.ShowImage = True
        '
        'ButtonPostPerSK
        '
        Me.ButtonPostPerSK.Image = CType(resources.GetObject("ButtonPostPerSK.Image"), System.Drawing.Image)
        Me.ButtonPostPerSK.Label = "Post Per SK"
        Me.ButtonPostPerSK.Name = "ButtonPostPerSK"
        Me.ButtonPostPerSK.ScreenTip = "Post Statistical Keyfigures"
        Me.ButtonPostPerSK.ShowImage = True
        '
        'SAPPsPlan
        '
        Me.SAPPsPlan.Items.Add(Me.ButtonPsUpdCheck)
        Me.SAPPsPlan.Items.Add(Me.ButtonPsUpdPost)
        Me.SAPPsPlan.Label = "PS Plan"
        Me.SAPPsPlan.Name = "SAPPsPlan"
        '
        'ButtonPsUpdCheck
        '
        Me.ButtonPsUpdCheck.Image = CType(resources.GetObject("ButtonPsUpdCheck.Image"), System.Drawing.Image)
        Me.ButtonPsUpdCheck.Label = "Check ERP Update"
        Me.ButtonPsUpdCheck.Name = "ButtonPsUpdCheck"
        Me.ButtonPsUpdCheck.ScreenTip = "Check ERP PS Plan Data"
        Me.ButtonPsUpdCheck.ShowImage = True
        '
        'ButtonPsUpdPost
        '
        Me.ButtonPsUpdPost.Image = CType(resources.GetObject("ButtonPsUpdPost.Image"), System.Drawing.Image)
        Me.ButtonPsUpdPost.Label = "Post ERP Update"
        Me.ButtonPsUpdPost.Name = "ButtonPsUpdPost"
        Me.ButtonPsUpdPost.ScreenTip = "Post ERP PS Plan Data"
        Me.ButtonPsUpdPost.ShowImage = True
        '
        'SAPCoCosting
        '
        Me.SAPCoCosting.Items.Add(Me.ButtonCostingCreate)
        Me.SAPCoCosting.Items.Add(Me.ButtonCostingChange)
        Me.SAPCoCosting.Label = "CO-PC Costing"
        Me.SAPCoCosting.Name = "SAPCoCosting"
        '
        'ButtonCostingCreate
        '
        Me.ButtonCostingCreate.Image = CType(resources.GetObject("ButtonCostingCreate.Image"), System.Drawing.Image)
        Me.ButtonCostingCreate.Label = "Create Costing"
        Me.ButtonCostingCreate.Name = "ButtonCostingCreate"
        Me.ButtonCostingCreate.ScreenTip = "Create CO-PC Costing"
        Me.ButtonCostingCreate.ShowImage = True
        '
        'ButtonCostingChange
        '
        Me.ButtonCostingChange.Image = CType(resources.GetObject("ButtonCostingChange.Image"), System.Drawing.Image)
        Me.ButtonCostingChange.Label = "Change Costing"
        Me.ButtonCostingChange.Name = "ButtonCostingChange"
        Me.ButtonCostingChange.ScreenTip = "Change CO-PC Costing"
        Me.ButtonCostingChange.ShowImage = True
        '
        'SAPCoPlLogon
        '
        Me.SAPCoPlLogon.Items.Add(Me.ButtonLogon)
        Me.SAPCoPlLogon.Items.Add(Me.ButtonLogoff)
        Me.SAPCoPlLogon.Label = "Logon"
        Me.SAPCoPlLogon.Name = "SAPCoPlLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapCoPlRibbon
        '
        Me.Name = "SapCoPlRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapCoPl)
        Me.SapCoPl.ResumeLayout(False)
        Me.SapCoPl.PerformLayout()
        Me.SAPCoOmPlan.ResumeLayout(False)
        Me.SAPCoOmPlan.PerformLayout()
        Me.SAPCoOmPlanPer.ResumeLayout(False)
        Me.SAPCoOmPlanPer.PerformLayout()
        Me.SAPPsPlan.ResumeLayout(False)
        Me.SAPPsPlan.PerformLayout()
        Me.SAPCoCosting.ResumeLayout(False)
        Me.SAPCoCosting.PerformLayout()
        Me.SAPCoPlLogon.ResumeLayout(False)
        Me.SAPCoPlLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCoPl As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPCoOmPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonReadAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoOmPlanPer As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonReadPerAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPerSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonReadPerSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoPlLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPPsPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPsUpdCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPsUpdPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoCosting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCostingCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCostingChange As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapCoPlRibbon
        Get
            Return Me.GetRibbon(Of SapCoPlRibbon)()
        End Get
    End Property
End Class
