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
        Me.Box3 = Me.Factory.CreateRibbonBox
        Me.ButtonCheckAO = Me.Factory.CreateRibbonButton
        Me.ButtonPostAO = Me.Factory.CreateRibbonButton
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.ButtonCheckPC = Me.Factory.CreateRibbonButton
        Me.ButtonPostPC = Me.Factory.CreateRibbonButton
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.ButtonCheckAI = Me.Factory.CreateRibbonButton
        Me.ButtonPostAI = Me.Factory.CreateRibbonButton
        Me.Box4 = Me.Factory.CreateRibbonBox
        Me.ButtonCheckSK = Me.Factory.CreateRibbonButton
        Me.ButtonPostSK = Me.Factory.CreateRibbonButton
        Me.SAPPsPlan = Me.Factory.CreateRibbonGroup
        Me.ButtonPsUpdCheck = Me.Factory.CreateRibbonButton
        Me.ButtonPsUpdPost = Me.Factory.CreateRibbonButton
        Me.SAPCoCosting = Me.Factory.CreateRibbonGroup
        Me.ButtonCostingCreate = Me.Factory.CreateRibbonButton
        Me.ButtonCostingChange = Me.Factory.CreateRibbonButton
        Me.SapPsMdGenerate = Me.Factory.CreateRibbonGroup
        Me.ButtonGenData = Me.Factory.CreateRibbonButton
        Me.SAPCoPlLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapCoPl.SuspendLayout()
        Me.SAPCoOmPlan.SuspendLayout()
        Me.Box3.SuspendLayout()
        Me.Box2.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Box4.SuspendLayout()
        Me.SAPPsPlan.SuspendLayout()
        Me.SAPCoCosting.SuspendLayout()
        Me.SapPsMdGenerate.SuspendLayout()
        Me.SAPCoPlLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapCoPl
        '
        Me.SapCoPl.Groups.Add(Me.SAPCoOmPlan)
        Me.SapCoPl.Groups.Add(Me.SAPPsPlan)
        Me.SapCoPl.Groups.Add(Me.SAPCoCosting)
        Me.SapCoPl.Groups.Add(Me.SapPsMdGenerate)
        Me.SapCoPl.Groups.Add(Me.SAPCoPlLogon)
        Me.SapCoPl.Label = "SAP CO-PL"
        Me.SapCoPl.Name = "SapCoPl"
        Me.SapCoPl.Position = Me.Factory.RibbonPosition.AfterOfficeId("SapCo")
        '
        'SAPCoOmPlan
        '
        Me.SAPCoOmPlan.Items.Add(Me.Box3)
        Me.SAPCoOmPlan.Items.Add(Me.Box2)
        Me.SAPCoOmPlan.Items.Add(Me.Box1)
        Me.SAPCoOmPlan.Items.Add(Me.Box4)
        Me.SAPCoOmPlan.Label = "CO-OM Plan"
        Me.SAPCoOmPlan.Name = "SAPCoOmPlan"
        '
        'Box3
        '
        Me.Box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box3.Items.Add(Me.ButtonCheckAO)
        Me.Box3.Items.Add(Me.ButtonPostAO)
        Me.Box3.Name = "Box3"
        '
        'ButtonCheckAO
        '
        Me.ButtonCheckAO.Image = CType(resources.GetObject("ButtonCheckAO.Image"), System.Drawing.Image)
        Me.ButtonCheckAO.Label = "Check AO"
        Me.ButtonCheckAO.Name = "ButtonCheckAO"
        Me.ButtonCheckAO.ScreenTip = "Check Activity Output"
        Me.ButtonCheckAO.ShowImage = True
        '
        'ButtonPostAO
        '
        Me.ButtonPostAO.Image = CType(resources.GetObject("ButtonPostAO.Image"), System.Drawing.Image)
        Me.ButtonPostAO.Label = "Post AO"
        Me.ButtonPostAO.Name = "ButtonPostAO"
        Me.ButtonPostAO.ScreenTip = "Post Activity Output"
        Me.ButtonPostAO.ShowImage = True
        '
        'Box2
        '
        Me.Box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box2.Items.Add(Me.ButtonCheckPC)
        Me.Box2.Items.Add(Me.ButtonPostPC)
        Me.Box2.Name = "Box2"
        '
        'ButtonCheckPC
        '
        Me.ButtonCheckPC.Image = CType(resources.GetObject("ButtonCheckPC.Image"), System.Drawing.Image)
        Me.ButtonCheckPC.Label = "Check PC"
        Me.ButtonCheckPC.Name = "ButtonCheckPC"
        Me.ButtonCheckPC.ScreenTip = "Check Primary Cost"
        Me.ButtonCheckPC.ShowImage = True
        '
        'ButtonPostPC
        '
        Me.ButtonPostPC.Image = CType(resources.GetObject("ButtonPostPC.Image"), System.Drawing.Image)
        Me.ButtonPostPC.Label = "Post PC"
        Me.ButtonPostPC.Name = "ButtonPostPC"
        Me.ButtonPostPC.ScreenTip = "Post Primary Cost"
        Me.ButtonPostPC.ShowImage = True
        '
        'Box1
        '
        Me.Box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box1.Items.Add(Me.ButtonCheckAI)
        Me.Box1.Items.Add(Me.ButtonPostAI)
        Me.Box1.Name = "Box1"
        '
        'ButtonCheckAI
        '
        Me.ButtonCheckAI.Image = CType(resources.GetObject("ButtonCheckAI.Image"), System.Drawing.Image)
        Me.ButtonCheckAI.Label = "Check AI"
        Me.ButtonCheckAI.Name = "ButtonCheckAI"
        Me.ButtonCheckAI.ScreenTip = "Check Activity Input"
        Me.ButtonCheckAI.ShowImage = True
        '
        'ButtonPostAI
        '
        Me.ButtonPostAI.Image = CType(resources.GetObject("ButtonPostAI.Image"), System.Drawing.Image)
        Me.ButtonPostAI.Label = "Post AI"
        Me.ButtonPostAI.Name = "ButtonPostAI"
        Me.ButtonPostAI.ScreenTip = "Post Activity Input"
        Me.ButtonPostAI.ShowImage = True
        '
        'Box4
        '
        Me.Box4.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box4.Items.Add(Me.ButtonCheckSK)
        Me.Box4.Items.Add(Me.ButtonPostSK)
        Me.Box4.Name = "Box4"
        '
        'ButtonCheckSK
        '
        Me.ButtonCheckSK.Image = CType(resources.GetObject("ButtonCheckSK.Image"), System.Drawing.Image)
        Me.ButtonCheckSK.Label = "Check SK"
        Me.ButtonCheckSK.Name = "ButtonCheckSK"
        Me.ButtonCheckSK.ScreenTip = "Check Statistical Keyfigures"
        Me.ButtonCheckSK.ShowImage = True
        '
        'ButtonPostSK
        '
        Me.ButtonPostSK.Image = CType(resources.GetObject("ButtonPostSK.Image"), System.Drawing.Image)
        Me.ButtonPostSK.Label = "Post SK"
        Me.ButtonPostSK.Name = "ButtonPostSK"
        Me.ButtonPostSK.ScreenTip = "Post Statistical Keyfigures"
        Me.ButtonPostSK.ShowImage = True
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
        'SapPsMdGenerate
        '
        Me.SapPsMdGenerate.Items.Add(Me.ButtonGenData)
        Me.SapPsMdGenerate.Label = "Generate"
        Me.SapPsMdGenerate.Name = "SapPsMdGenerate"
        '
        'ButtonGenData
        '
        Me.ButtonGenData.Image = CType(resources.GetObject("ButtonGenData.Image"), System.Drawing.Image)
        Me.ButtonGenData.Label = "Generate Data"
        Me.ButtonGenData.Name = "ButtonGenData"
        Me.ButtonGenData.ScreenTip = "Generate the Output Data"
        Me.ButtonGenData.ShowImage = True
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
        Me.Box3.ResumeLayout(False)
        Me.Box3.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Box4.ResumeLayout(False)
        Me.Box4.PerformLayout()
        Me.SAPPsPlan.ResumeLayout(False)
        Me.SAPPsPlan.PerformLayout()
        Me.SAPCoCosting.ResumeLayout(False)
        Me.SAPCoCosting.PerformLayout()
        Me.SapPsMdGenerate.ResumeLayout(False)
        Me.SapPsMdGenerate.PerformLayout()
        Me.SAPCoPlLogon.ResumeLayout(False)
        Me.SAPCoPlLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapCoPl As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPCoOmPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCheckAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCheckPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCheckAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAO As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostPC As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostAI As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPostSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCheckSK As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoPlLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPPsPlan As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonPsUpdCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonPsUpdPost As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SAPCoCosting As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonCostingCreate As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonCostingChange As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Box3 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents Box4 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents SapPsMdGenerate As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonGenData As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Ribbon1() As SapCoPlRibbon
        Get
            Return Me.GetRibbon(Of SapCoPlRibbon)()
        End Get
    End Property
End Class
