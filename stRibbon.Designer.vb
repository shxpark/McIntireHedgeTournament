Partial Class stRibbon
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
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.AlphaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.BetaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.GammaTBtn = Me.Factory.CreateRibbonToggleButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.ManualTBtn = Me.Factory.CreateRibbonToggleButton
        Me.SynchTBtn = Me.Factory.CreateRibbonToggleButton
        Me.SimTBtn = Me.Factory.CreateRibbonToggleButton
        Me.AutoTBtn = Me.Factory.CreateRibbonToggleButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.DashboardBtn = Me.Factory.CreateRibbonButton
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.InitialPositionsBtn = Me.Factory.CreateRibbonButton
        Me.AcquiredPositionsBtn = Me.Factory.CreateRibbonButton
        Me.TransactionQBtn = Me.Factory.CreateRibbonButton
        Me.ResetAPBtn = Me.Factory.CreateRibbonButton
        Me.UploadAPBtn = Me.Factory.CreateRibbonButton
        Me.EditAPBtn = Me.Factory.CreateRibbonButton
        Me.ConfirmationBtn = Me.Factory.CreateRibbonButton
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.StockMktBtn = Me.Factory.CreateRibbonButton
        Me.OptionMktBtn = Me.Factory.CreateRibbonButton
        Me.SP500Btn = Me.Factory.CreateRibbonButton
        Me.SettingBtn = Me.Factory.CreateRibbonButton
        Me.TCostsBtn = Me.Factory.CreateRibbonButton
        Me.FinChartsBtn = Me.Factory.CreateRibbonButton
        Me.Group6 = Me.Factory.CreateRibbonGroup
        Me.QuitBtn = Me.Factory.CreateRibbonButton
        Me.IPOnOffBtn = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group6.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group6)
        Me.Tab1.Label = "Spartan Trader"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.AlphaTBtn)
        Me.Group1.Items.Add(Me.BetaTBtn)
        Me.Group1.Items.Add(Me.GammaTBtn)
        Me.Group1.Label = "Database"
        Me.Group1.Name = "Group1"
        '
        'AlphaTBtn
        '
        Me.AlphaTBtn.Label = "Alpha"
        Me.AlphaTBtn.Name = "AlphaTBtn"
        '
        'BetaTBtn
        '
        Me.BetaTBtn.Label = "Beta"
        Me.BetaTBtn.Name = "BetaTBtn"
        '
        'GammaTBtn
        '
        Me.GammaTBtn.Label = "Gamma"
        Me.GammaTBtn.Name = "GammaTBtn"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.ManualTBtn)
        Me.Group2.Items.Add(Me.SynchTBtn)
        Me.Group2.Items.Add(Me.SimTBtn)
        Me.Group2.Items.Add(Me.AutoTBtn)
        Me.Group2.Label = "Mode"
        Me.Group2.Name = "Group2"
        '
        'ManualTBtn
        '
        Me.ManualTBtn.Label = "Manual"
        Me.ManualTBtn.Name = "ManualTBtn"
        '
        'SynchTBtn
        '
        Me.SynchTBtn.Label = "Synch"
        Me.SynchTBtn.Name = "SynchTBtn"
        '
        'SimTBtn
        '
        Me.SimTBtn.Label = "Sim"
        Me.SimTBtn.Name = "SimTBtn"
        '
        'AutoTBtn
        '
        Me.AutoTBtn.Label = "Auto"
        Me.AutoTBtn.Name = "AutoTBtn"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.DashboardBtn)
        Me.Group3.Label = "Dashboard"
        Me.Group3.Name = "Group3"
        '
        'DashboardBtn
        '
        Me.DashboardBtn.Label = "Dashboard"
        Me.DashboardBtn.Name = "DashboardBtn"
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.InitialPositionsBtn)
        Me.Group4.Items.Add(Me.AcquiredPositionsBtn)
        Me.Group4.Items.Add(Me.TransactionQBtn)
        Me.Group4.Items.Add(Me.ResetAPBtn)
        Me.Group4.Items.Add(Me.UploadAPBtn)
        Me.Group4.Items.Add(Me.EditAPBtn)
        Me.Group4.Items.Add(Me.ConfirmationBtn)
        Me.Group4.Label = "Portfolio Management"
        Me.Group4.Name = "Group4"
        '
        'InitialPositionsBtn
        '
        Me.InitialPositionsBtn.Label = "Initial Positions"
        Me.InitialPositionsBtn.Name = "InitialPositionsBtn"
        '
        'AcquiredPositionsBtn
        '
        Me.AcquiredPositionsBtn.Label = "Acquired Positions"
        Me.AcquiredPositionsBtn.Name = "AcquiredPositionsBtn"
        '
        'TransactionQBtn
        '
        Me.TransactionQBtn.Label = "Transactions Q"
        Me.TransactionQBtn.Name = "TransactionQBtn"
        '
        'ResetAPBtn
        '
        Me.ResetAPBtn.Label = "Reset AP"
        Me.ResetAPBtn.Name = "ResetAPBtn"
        '
        'UploadAPBtn
        '
        Me.UploadAPBtn.Label = "Upload AP"
        Me.UploadAPBtn.Name = "UploadAPBtn"
        '
        'EditAPBtn
        '
        Me.EditAPBtn.Label = "Edit AP"
        Me.EditAPBtn.Name = "EditAPBtn"
        '
        'ConfirmationBtn
        '
        Me.ConfirmationBtn.Label = "Confirmation"
        Me.ConfirmationBtn.Name = "ConfirmationBtn"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.StockMktBtn)
        Me.Group5.Items.Add(Me.OptionMktBtn)
        Me.Group5.Items.Add(Me.SP500Btn)
        Me.Group5.Items.Add(Me.SettingBtn)
        Me.Group5.Items.Add(Me.TCostsBtn)
        Me.Group5.Items.Add(Me.FinChartsBtn)
        Me.Group5.Label = "Business Intelligence"
        Me.Group5.Name = "Group5"
        '
        'StockMktBtn
        '
        Me.StockMktBtn.Label = "Stock Mkt"
        Me.StockMktBtn.Name = "StockMktBtn"
        '
        'OptionMktBtn
        '
        Me.OptionMktBtn.Label = "Option Mkt"
        Me.OptionMktBtn.Name = "OptionMktBtn"
        '
        'SP500Btn
        '
        Me.SP500Btn.Label = "SP 500"
        Me.SP500Btn.Name = "SP500Btn"
        '
        'SettingBtn
        '
        Me.SettingBtn.Label = "Settings"
        Me.SettingBtn.Name = "SettingBtn"
        '
        'TCostsBtn
        '
        Me.TCostsBtn.Label = "T Costs"
        Me.TCostsBtn.Name = "TCostsBtn"
        '
        'FinChartsBtn
        '
        Me.FinChartsBtn.Label = "Fin Charts"
        Me.FinChartsBtn.Name = "FinChartsBtn"
        '
        'Group6
        '
        Me.Group6.Items.Add(Me.QuitBtn)
        Me.Group6.Items.Add(Me.IPOnOffBtn)
        Me.Group6.Label = "Control"
        Me.Group6.Name = "Group6"
        '
        'QuitBtn
        '
        Me.QuitBtn.Label = "Quit"
        Me.QuitBtn.Name = "QuitBtn"
        '
        'IPOnOffBtn
        '
        Me.IPOnOffBtn.Label = "IP On/Off"
        Me.IPOnOffBtn.Name = "IPOnOffBtn"
        '
        'stRibbon
        '
        Me.Name = "stRibbon"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group6.ResumeLayout(False)
        Me.Group6.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents AlphaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents BetaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents GammaTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ManualTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents DashboardBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents InitialPositionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents AcquiredPositionsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TransactionQBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ConfirmationBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents StockMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents OptionMktBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SP500Btn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SettingBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents TCostsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group6 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents QuitBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ResetAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents UploadAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents EditAPBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents FinChartsBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents IPOnOffBtn As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents SynchTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents SimTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
    Friend WithEvents AutoTBtn As Microsoft.Office.Tools.Ribbon.RibbonToggleButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property stRibbon() As stRibbon
        Get
            Return Me.GetRibbon(Of stRibbon)()
        End Get
    End Property
End Class
