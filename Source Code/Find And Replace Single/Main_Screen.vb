Imports System.IO
Imports Microsoft.Win32

Public Class Main_Screen
    Inherits System.Windows.Forms.Form

    Dim WithEvents Worker1 As Worker

    Private workerbusy As Boolean = False
    Private steps As Integer = 0
    'steps 0: process not launched
    'steps 1: folder count
    'steps 2: file count (and file queue creation)
    'steps 3: file search and replace operation
    'steps 4: folder renaming

    Private thread1snapshot As Long


    Private queueselector As Integer = 1
    Private filequeue1 As Queue

    Public Delegate Sub WorkerComplete_h()
    Public Delegate Sub WorkerError_h(ByVal Message As Exception)
    Public Delegate Sub WorkerFolderCount_h(ByVal Result As Long)
    Public Delegate Sub WorkerFolderRename_h(ByVal Result As Long)
    Public Delegate Sub WorkerFileCount_h(ByVal Result As Long)
    Public Delegate Sub WorkerStatusMessage_h(ByVal message As String, ByVal statustag As Integer)
    Public Delegate Sub WorkerFileProcessing_h(ByVal filename As String, ByVal queue As Integer)
    Public Delegate Sub WorkerStepAnnounce_h(ByVal stepnumber As String)

    Public Delegate Sub WorkerContentChanges_h()
    Public Delegate Sub WorkerFileRenames_h()

    Private filetypes As ArrayList

    Private frmsearchstring, frmreplacestring, frmbasefolder As String

    Public dataloaded As Boolean = False



#Region " Windows Form Designer generated code "

    

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Worker1 = New Worker
        AddHandler Worker1.WorkerComplete, AddressOf WorkerCompleteHandler
        AddHandler Worker1.WorkerError, AddressOf WorkerErrorHandler
        AddHandler Worker1.WorkerFolderCount, AddressOf WorkerFolderCountHandler
        AddHandler Worker1.WorkerFolderRename, AddressOf WorkerFolderRenameHandler
        AddHandler Worker1.WorkerFileCount, AddressOf WorkerFileCountHandler
        AddHandler Worker1.WorkerStatusMessage, AddressOf WorkerStatusMessageHandler
        AddHandler Worker1.WorkerFileProcessing, AddressOf WorkerFileProcessingHandler
        AddHandler Worker1.WorkerStepAnnounce, AddressOf WorkerStepAnnounceHandler
        AddHandler Worker1.WorkerContentChanges, AddressOf WorkerContentChangesHandler
        AddHandler Worker1.WorkerFileRenames, AddressOf WorkerFileRenamesHandler

        filequeue1 = New Queue
      

        filetypes = New ArrayList


    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents txtSearchString As System.Windows.Forms.TextBox
    Friend WithEvents txtReplaceString As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ButtonFolderBrowse As System.Windows.Forms.Button
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents txtBaseFolder As System.Windows.Forms.TextBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents ButtonOperationLaunch As System.Windows.Forms.Button
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents step1 As System.Windows.Forms.Label
    Friend WithEvents step2 As System.Windows.Forms.Label
    Friend WithEvents step3 As System.Windows.Forms.Label
    Friend WithEvents step4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtFolderRenames As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents txtFileQueueTotal1 As System.Windows.Forms.Label
    Friend WithEvents txtFileQueue1 As System.Windows.Forms.Label
    Friend WithEvents txtStatusBar1 As System.Windows.Forms.TextBox
    Friend WithEvents txtFileCount As System.Windows.Forms.Label
    Friend WithEvents txtFolderCount As System.Windows.Forms.Label
    Friend WithEvents txtProcessLaunched As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents txtContentChanges As System.Windows.Forms.Label
    Friend WithEvents txtFileRenames As System.Windows.Forms.Label
    Friend WithEvents txtProcessEnded As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Button2 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Main_Screen))
        Me.txtSearchString = New System.Windows.Forms.TextBox
        Me.txtReplaceString = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.txtBaseFolder = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.ButtonFolderBrowse = New System.Windows.Forms.Button
        Me.Label4 = New System.Windows.Forms.Label
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.ButtonOperationLaunch = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.step1 = New System.Windows.Forms.Label
        Me.step2 = New System.Windows.Forms.Label
        Me.step3 = New System.Windows.Forms.Label
        Me.step4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.txtFolderRenames = New System.Windows.Forms.Label
        Me.Label12 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.txtFileQueueTotal1 = New System.Windows.Forms.Label
        Me.txtFileQueue1 = New System.Windows.Forms.Label
        Me.txtStatusBar1 = New System.Windows.Forms.TextBox
        Me.txtFileCount = New System.Windows.Forms.Label
        Me.txtFolderCount = New System.Windows.Forms.Label
        Me.txtProcessLaunched = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label13 = New System.Windows.Forms.Label
        Me.txtProcessEnded = New System.Windows.Forms.Label
        Me.txtContentChanges = New System.Windows.Forms.Label
        Me.txtFileRenames = New System.Windows.Forms.Label
        Me.Label14 = New System.Windows.Forms.Label
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtSearchString
        '
        Me.txtSearchString.BackColor = System.Drawing.Color.White
        Me.txtSearchString.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtSearchString.ForeColor = System.Drawing.Color.Black
        Me.txtSearchString.Location = New System.Drawing.Point(8, 56)
        Me.txtSearchString.Name = "txtSearchString"
        Me.txtSearchString.Size = New System.Drawing.Size(288, 20)
        Me.txtSearchString.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.txtSearchString, "The search string searched for during the Find And Replace operation")
        '
        'txtReplaceString
        '
        Me.txtReplaceString.BackColor = System.Drawing.Color.White
        Me.txtReplaceString.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtReplaceString.ForeColor = System.Drawing.Color.Black
        Me.txtReplaceString.Location = New System.Drawing.Point(8, 96)
        Me.txtReplaceString.Name = "txtReplaceString"
        Me.txtReplaceString.Size = New System.Drawing.Size(288, 20)
        Me.txtReplaceString.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtReplaceString, "The replace string used during the Find And Replace operation")
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Black
        Me.Label1.Location = New System.Drawing.Point(8, 80)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "SEARCH STRING"
        '
        'txtBaseFolder
        '
        Me.txtBaseFolder.BackColor = System.Drawing.Color.White
        Me.txtBaseFolder.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.txtBaseFolder.ForeColor = System.Drawing.Color.Black
        Me.txtBaseFolder.Location = New System.Drawing.Point(8, 136)
        Me.txtBaseFolder.Name = "txtBaseFolder"
        Me.txtBaseFolder.ReadOnly = True
        Me.txtBaseFolder.Size = New System.Drawing.Size(224, 20)
        Me.txtBaseFolder.TabIndex = 4
        Me.ToolTip1.SetToolTip(Me.txtBaseFolder, "The base folder from which the search is initiated during the Find And Replace op" & _
                "eration")
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Black
        Me.Label2.Location = New System.Drawing.Point(8, 120)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(100, 16)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "REPLACE STRING"
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Black
        Me.Label3.Location = New System.Drawing.Point(8, 16)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(280, 32)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "ENTER THE DESIRED SEARCH AND REPLACE STRINGS, SELECT THE OPERATION BASE FOLDER AN" & _
            "D CLICK ON THE PROCESS BUTTON TO LAUNCH THE FIND AND REPLACE OPERATION.  (DELIMI" & _
            "TER: ';;')"
        '
        'ButtonFolderBrowse
        '
        Me.ButtonFolderBrowse.BackColor = System.Drawing.Color.Silver
        Me.ButtonFolderBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonFolderBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonFolderBrowse.Location = New System.Drawing.Point(240, 136)
        Me.ButtonFolderBrowse.Name = "ButtonFolderBrowse"
        Me.ButtonFolderBrowse.Size = New System.Drawing.Size(56, 20)
        Me.ButtonFolderBrowse.TabIndex = 7
        Me.ButtonFolderBrowse.Text = "BROWSE"
        Me.ToolTip1.SetToolTip(Me.ButtonFolderBrowse, "Launches the Folder Browser Dialog")
        Me.ButtonFolderBrowse.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.Black
        Me.Label4.Location = New System.Drawing.Point(8, 160)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(136, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "BASE FOLDER"
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select the base folder from which the search is initiated during the Find And Rep" & _
            "lace operation"
        Me.FolderBrowserDialog1.ShowNewFolderButton = False
        '
        'ButtonOperationLaunch
        '
        Me.ButtonOperationLaunch.BackColor = System.Drawing.Color.Silver
        Me.ButtonOperationLaunch.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ButtonOperationLaunch.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonOperationLaunch.ForeColor = System.Drawing.Color.Black
        Me.ButtonOperationLaunch.Location = New System.Drawing.Point(104, 182)
        Me.ButtonOperationLaunch.Name = "ButtonOperationLaunch"
        Me.ButtonOperationLaunch.Size = New System.Drawing.Size(88, 20)
        Me.ButtonOperationLaunch.TabIndex = 10
        Me.ButtonOperationLaunch.Text = "Process"
        Me.ToolTip1.SetToolTip(Me.ButtonOperationLaunch, "Launches Find and Replace Operation")
        Me.ButtonOperationLaunch.UseVisualStyleBackColor = False
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.Silver
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(640, 176)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 18)
        Me.Button1.TabIndex = 37
        Me.Button1.Text = "KILL PROCESSES"
        Me.ToolTip1.SetToolTip(Me.Button1, "Launches the Folder Browser Dialog")
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Button2
        '
        Me.Button2.BackColor = System.Drawing.Color.Orchid
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.Button2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(264, 368)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(15, 13)
        Me.Button2.TabIndex = 51
        Me.ToolTip1.SetToolTip(Me.Button2, "Force Application Status Check")
        Me.Button2.UseVisualStyleBackColor = False
        '
        'txtStatus
        '
        Me.txtStatus.BackColor = System.Drawing.Color.Plum
        Me.txtStatus.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatus.ForeColor = System.Drawing.Color.White
        Me.txtStatus.Location = New System.Drawing.Point(320, 368)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.ReadOnly = True
        Me.txtStatus.Size = New System.Drawing.Size(416, 13)
        Me.txtStatus.TabIndex = 15
        Me.txtStatus.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 5000
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.ListBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListBox1.ForeColor = System.Drawing.Color.DimGray
        Me.ListBox1.Location = New System.Drawing.Point(640, 40)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(88, 132)
        Me.ListBox1.TabIndex = 31
        '
        'step1
        '
        Me.step1.ForeColor = System.Drawing.Color.Gray
        Me.step1.Location = New System.Drawing.Point(38, 224)
        Me.step1.Name = "step1"
        Me.step1.Size = New System.Drawing.Size(184, 16)
        Me.step1.TabIndex = 32
        Me.step1.Text = "Step 1: Retrieve Folder List"
        '
        'step2
        '
        Me.step2.ForeColor = System.Drawing.Color.Gray
        Me.step2.Location = New System.Drawing.Point(188, 224)
        Me.step2.Name = "step2"
        Me.step2.Size = New System.Drawing.Size(184, 16)
        Me.step2.TabIndex = 33
        Me.step2.Text = "Step 2: Retrieve File List"
        '
        'step3
        '
        Me.step3.ForeColor = System.Drawing.Color.Gray
        Me.step3.Location = New System.Drawing.Point(334, 224)
        Me.step3.Name = "step3"
        Me.step3.Size = New System.Drawing.Size(216, 16)
        Me.step3.TabIndex = 34
        Me.step3.Text = "Step 3: Process File Names and Content"
        '
        'step4
        '
        Me.step4.ForeColor = System.Drawing.Color.Gray
        Me.step4.Location = New System.Drawing.Point(547, 224)
        Me.step4.Name = "step4"
        Me.step4.Size = New System.Drawing.Size(184, 16)
        Me.step4.TabIndex = 35
        Me.step4.Text = "Step 4: Process Folder Names"
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Label5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label5.Location = New System.Drawing.Point(8, 176)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(288, 32)
        Me.Label5.TabIndex = 36
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.Color.Black
        Me.Label6.Location = New System.Drawing.Point(640, 16)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(104, 16)
        Me.Label6.TabIndex = 38
        Me.Label6.Text = "FILE TYPES TO BE SCANNED"
        '
        'txtFolderRenames
        '
        Me.txtFolderRenames.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFolderRenames.ForeColor = System.Drawing.Color.Gray
        Me.txtFolderRenames.Location = New System.Drawing.Point(7, 127)
        Me.txtFolderRenames.Name = "txtFolderRenames"
        Me.txtFolderRenames.Size = New System.Drawing.Size(216, 16)
        Me.txtFolderRenames.TabIndex = 45
        Me.txtFolderRenames.Text = "Folder Renames:"
        Me.txtFolderRenames.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Label12.ForeColor = System.Drawing.Color.Gray
        Me.Label12.Location = New System.Drawing.Point(320, 112)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(192, 16)
        Me.Label12.TabIndex = 44
        Me.Label12.Text = "File Scan Progress:"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Label7.ForeColor = System.Drawing.Color.Gray
        Me.Label7.Location = New System.Drawing.Point(336, 128)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(64, 16)
        Me.Label7.TabIndex = 39
        Me.Label7.Text = "Worker:"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileQueueTotal1
        '
        Me.txtFileQueueTotal1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFileQueueTotal1.ForeColor = System.Drawing.Color.Gray
        Me.txtFileQueueTotal1.Location = New System.Drawing.Point(440, 128)
        Me.txtFileQueueTotal1.Name = "txtFileQueueTotal1"
        Me.txtFileQueueTotal1.Size = New System.Drawing.Size(80, 16)
        Me.txtFileQueueTotal1.TabIndex = 26
        Me.txtFileQueueTotal1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFileQueue1
        '
        Me.txtFileQueue1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFileQueue1.ForeColor = System.Drawing.Color.Gray
        Me.txtFileQueue1.Location = New System.Drawing.Point(400, 128)
        Me.txtFileQueue1.Name = "txtFileQueue1"
        Me.txtFileQueue1.Size = New System.Drawing.Size(40, 16)
        Me.txtFileQueue1.TabIndex = 21
        Me.txtFileQueue1.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtStatusBar1
        '
        Me.txtStatusBar1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtStatusBar1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtStatusBar1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtStatusBar1.ForeColor = System.Drawing.Color.Gray
        Me.txtStatusBar1.Location = New System.Drawing.Point(7, 170)
        Me.txtStatusBar1.Name = "txtStatusBar1"
        Me.txtStatusBar1.ReadOnly = True
        Me.txtStatusBar1.Size = New System.Drawing.Size(296, 10)
        Me.txtStatusBar1.TabIndex = 14
        '
        'txtFileCount
        '
        Me.txtFileCount.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFileCount.ForeColor = System.Drawing.Color.Gray
        Me.txtFileCount.Location = New System.Drawing.Point(320, 96)
        Me.txtFileCount.Name = "txtFileCount"
        Me.txtFileCount.Size = New System.Drawing.Size(192, 16)
        Me.txtFileCount.TabIndex = 16
        Me.txtFileCount.Text = "File Count:"
        Me.txtFileCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFolderCount
        '
        Me.txtFolderCount.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFolderCount.ForeColor = System.Drawing.Color.Gray
        Me.txtFolderCount.Location = New System.Drawing.Point(320, 80)
        Me.txtFolderCount.Name = "txtFolderCount"
        Me.txtFolderCount.Size = New System.Drawing.Size(192, 16)
        Me.txtFolderCount.TabIndex = 13
        Me.txtFolderCount.Text = "Folder Count:"
        Me.txtFolderCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtProcessLaunched
        '
        Me.txtProcessLaunched.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtProcessLaunched.ForeColor = System.Drawing.Color.Gray
        Me.txtProcessLaunched.Location = New System.Drawing.Point(320, 64)
        Me.txtProcessLaunched.Name = "txtProcessLaunched"
        Me.txtProcessLaunched.Size = New System.Drawing.Size(192, 16)
        Me.txtProcessLaunched.TabIndex = 12
        Me.txtProcessLaunched.Text = "Launched:"
        Me.txtProcessLaunched.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.Label13)
        Me.Panel1.Controls.Add(Me.txtFolderRenames)
        Me.Panel1.Controls.Add(Me.txtProcessEnded)
        Me.Panel1.Controls.Add(Me.txtStatusBar1)
        Me.Panel1.ForeColor = System.Drawing.Color.Gray
        Me.Panel1.Location = New System.Drawing.Point(312, 16)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(312, 192)
        Me.Panel1.TabIndex = 46
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Thistle
        Me.Label13.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label13.ForeColor = System.Drawing.Color.Black
        Me.Label13.Location = New System.Drawing.Point(8, 8)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(296, 24)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Operation Details"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'txtProcessEnded
        '
        Me.txtProcessEnded.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtProcessEnded.ForeColor = System.Drawing.Color.Gray
        Me.txtProcessEnded.Location = New System.Drawing.Point(7, 143)
        Me.txtProcessEnded.Name = "txtProcessEnded"
        Me.txtProcessEnded.Size = New System.Drawing.Size(192, 16)
        Me.txtProcessEnded.TabIndex = 13
        Me.txtProcessEnded.Text = "Ended:"
        Me.txtProcessEnded.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtContentChanges
        '
        Me.txtContentChanges.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtContentChanges.ForeColor = System.Drawing.Color.Gray
        Me.txtContentChanges.Location = New System.Drawing.Point(530, 150)
        Me.txtContentChanges.Name = "txtContentChanges"
        Me.txtContentChanges.Size = New System.Drawing.Size(62, 16)
        Me.txtContentChanges.TabIndex = 48
        Me.txtContentChanges.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtFileRenames
        '
        Me.txtFileRenames.BackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        Me.txtFileRenames.ForeColor = System.Drawing.Color.Gray
        Me.txtFileRenames.Location = New System.Drawing.Point(536, 168)
        Me.txtFileRenames.Name = "txtFileRenames"
        Me.txtFileRenames.Size = New System.Drawing.Size(62, 16)
        Me.txtFileRenames.TabIndex = 49
        Me.txtFileRenames.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label14
        '
        Me.Label14.ForeColor = System.Drawing.Color.Thistle
        Me.Label14.Location = New System.Drawing.Point(8, 368)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(256, 16)
        Me.Label14.TabIndex = 50
        '
        'Timer2
        '
        Me.Timer2.Enabled = True
        Me.Timer2.Interval = 500
        '
        'Main_Screen
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(740, 250)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.txtFileRenames)
        Me.Controls.Add(Me.txtContentChanges)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.step4)
        Me.Controls.Add(Me.step3)
        Me.Controls.Add(Me.step2)
        Me.Controls.Add(Me.step1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.txtFileQueueTotal1)
        Me.Controls.Add(Me.txtFileQueue1)
        Me.Controls.Add(Me.txtStatus)
        Me.Controls.Add(Me.txtBaseFolder)
        Me.Controls.Add(Me.txtReplaceString)
        Me.Controls.Add(Me.txtSearchString)
        Me.Controls.Add(Me.txtFileCount)
        Me.Controls.Add(Me.txtFolderCount)
        Me.Controls.Add(Me.txtProcessLaunched)
        Me.Controls.Add(Me.ButtonOperationLaunch)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.ButtonFolderBrowse)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Panel1)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(748, 284)
        Me.MinimumSize = New System.Drawing.Size(748, 284)
        Me.Name = "Main_Screen"
        Me.Text = "Find And Replace Single"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub Error_Handler(ByVal ex As Exception, Optional ByVal identifier_msg As String = "")
        Try
            If ex.Message.IndexOf("Thread was being aborted") < 0 Then
                If identifier_msg.Length > 0 Then
                    identifier_msg = identifier_msg & ": "
                End If

                Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg & ex.ToString)
                Display_Message1.ShowDialog()
                Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
                If dir.Exists = False Then
                    dir.Create()
                End If
                Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
                filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg & ex.ToString)
                filewriter.Flush()
                filewriter.Close()
            End If
        Catch exc As Exception

            MsgBox("An error occurred in Find And Replace's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub

    Private Sub Error_Handler(ByVal identifier_msg As String)
        Try

            Dim Display_Message1 As New Display_Message("The Application encountered the following problem: " & vbCrLf & identifier_msg)
            Display_Message1.ShowDialog()
            Dim dir As DirectoryInfo = New DirectoryInfo((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs")
            If dir.Exists = False Then
                dir.Create()
            End If
            Dim filewriter As StreamWriter = New StreamWriter((Application.StartupPath & "\").Replace("\\", "\") & "Error Logs\" & Format(Now(), "yyyyMMdd") & "_Error_Log.txt", True)
            filewriter.WriteLine("#" & Format(Now(), "dd/MM/yyyy HH:mm:ss") & " - " & identifier_msg)
            filewriter.Flush()
            filewriter.Close()

        Catch exc As Exception
            MsgBox("An error occurred in Find And Replace's error handling routine. The application will try to recover from this serious error.", MsgBoxStyle.Critical, "Critical Error Encountered")
        End Try
    End Sub


    Private Sub ButtonFolderBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFolderBrowse.Click
        Try
            Dim result As DialogResult
            If Not txtBaseFolder.Text = "" And txtBaseFolder.Text Is Nothing = False Then


                Dim foldercheck As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(txtBaseFolder.Text)
                If foldercheck.Exists = True Then
                    FolderBrowserDialog1.SelectedPath = txtBaseFolder.Text
                End If
            End If
            result = FolderBrowserDialog1.ShowDialog()
            If result = DialogResult.OK Or result = DialogResult.Yes Then
                txtBaseFolder.Text = FolderBrowserDialog1.SelectedPath
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub




    Private Sub SendMessage(ByVal labelname As String, ByVal message As String)
        Try
            Dim cont As Control
            'MsgBox(labelname & " -- " & message)
            For Each cont In Me.Controls
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
            'MsgBox(labelname & " -- " & message)
            For Each cont In Me.Panel1.Controls
                If cont.Name = labelname Then
                    cont.Text = message
                    cont.Refresh()
                    Exit For
                End If
            Next
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub ButtonOperationLaunch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOperationLaunch.Click

        SendMessage("txtProcessLaunched", "Launched:")
        SendMessage("txtProcessEnded", "Ended:")
        SendMessage("txtStatus", "")
        SendMessage("txtStatusBar1", "")
        SendMessage("txtStatusBar2", "")
        SendMessage("txtStatusBar3", "")
        SendMessage("txtStatusBar4", "")
        SendMessage("txtStatusBar5", "")
        SendMessage("txtFolderCount", "Folder Count:")
        SendMessage("txtFileCount", "File Count:")
        SendMessage("txtFileQueue1", "")
        SendMessage("txtFileQueueTotal1", "")
        SendMessage("txtFileQueue2", "")
        SendMessage("txtFileQueueTotal2", "")
        SendMessage("txtFileQueue3", "")
        SendMessage("txtFileQueueTotal3", "")
        SendMessage("txtFileQueue4", "")
        SendMessage("txtFileQueueTotal4", "")
        SendMessage("txtFileQueue5", "")
        SendMessage("txtFileQueueTotal5", "")
        SendMessage("txtFolderRenames", "Folder Renames:")

        filequeue1.Clear()


        SendMessage("txtProcessLaunched", "Launched: " & Format(Now(), "dd/MM/yyyy HH:mm:ss"))

        Worker1.filetypes.Clear()
        Dim str As String
        For Each str In filetypes
            Worker1.filetypes.Add(str)
        Next
        Worker1.basefolder = txtBaseFolder.Text.Trim
        Worker1.searchstring = txtSearchString.Text
        Worker1.replacestring = txtReplaceString.Text

        Dim delim() As String = {";;"}

        Dim search() As String = txtSearchString.Text.Split(delim, StringSplitOptions.None)
        Dim replace() As String = txtReplaceString.Text.Split(delim, StringSplitOptions.None)
        
        If search.Length = replace.Length Then
            Worker1.stopcheck = False

            steps = 1
            statuslabel()
            Worker1.ChooseThreads(1)
            workerbusy = True

            ButtonOperationLaunch.Enabled = False
        Else
            MsgBox("The number of search items does not match up to the number of replace items as delimited using the ';;' delimiter string. Please go back ensure that the number of terms match up.", MsgBoxStyle.Information, "Input Error")
        End If


    End Sub

    Public Sub WorkerStepAnnounceHandler(ByVal stepnumber As String)
        Try
            Select Case stepnumber
                Case "0"
                    step1.ForeColor = Color.Gray
                    step2.ForeColor = Color.Gray
                    step3.ForeColor = Color.Gray
                    step4.ForeColor = Color.Gray
                Case "1"
                    step1.ForeColor = Color.Black
                    step2.ForeColor = Color.Gray
                    step3.ForeColor = Color.Gray
                    step4.ForeColor = Color.Gray
                Case "2"
                    step1.ForeColor = Color.Gray
                    step2.ForeColor = Color.Black
                    step3.ForeColor = Color.Gray
                    step4.ForeColor = Color.Gray
                Case "3"
                    step1.ForeColor = Color.Gray
                    step2.ForeColor = Color.Gray
                    step3.ForeColor = Color.Black
                    step4.ForeColor = Color.Gray
                Case "4"
                    step1.ForeColor = Color.Gray
                    step2.ForeColor = Color.Gray
                    step3.ForeColor = Color.Gray
                    step4.ForeColor = Color.Black
                Case "5"
                    step1.ForeColor = Color.Gray
                    step2.ForeColor = Color.Gray
                    step3.ForeColor = Color.Gray
                    step4.ForeColor = Color.Gray
            End Select
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerStatusMessageHandler(ByVal message As String, ByVal statustag As Integer)
        Try
            If statustag = 1 Then
                SendMessage("txtStatus", message)
            Else
                SendMessage("txtStatusBar1", message.Replace(txtBaseFolder.Text & "\", "...\").ToUpper)
                If steps = 2 And Not message = "" Then
                    Select Case queueselector
                        Case 1
                            filequeue1.Enqueue(message.Replace("Examining: ", ""))

                    End Select
                    'queueselector = queueselector + 1
                    'If queueselector > 5 Then
                    '    queueselector = 1
                    'End If
                End If
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerErrorHandler(ByVal Message As Exception)
        Try
            Error_Handler(Message)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFolderCountHandler(ByVal Result As Long)
        Try
            SendMessage("txtFolderCount", "Folder Count: " & Result.ToString)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFolderRenameHandler(ByVal Result As Long, ByVal oftotal As Long)
        Try
            SendMessage("txtFolderRenames", "Folder Renames: " & Result.ToString & " (of " & oftotal.ToString & ")")

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileCountHandler(ByVal Result As Long)
        Try
            SendMessage("txtFileCount", "File Count: " & Result.ToString)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileRenamesHandler()
        Try
            Dim lng As Long
            If txtFileRenames.Text = "" Then
                lng = 0
            Else
                lng = CLng(txtFileRenames.Text)
            End If
            lng = lng + 1
            SendMessage("txtFileRenames", lng.ToString)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerContentChangesHandler()
        Try
            Dim lng As Long
            If txtContentChanges.Text = "" Then
                lng = 0
            Else
                lng = CLng(txtContentChanges.Text)
            End If
            lng = lng + 1
            SendMessage("txtContentChanges", lng.ToString)
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerCompleteHandler(ByVal queue As Integer)
        Try
            Dim eventhandled As Boolean = False
            workerbusy = False
            If steps = 1 And eventhandled = False Then
                steps = 2
                Worker1.ChooseThreads(2)
                workerbusy = True
                eventhandled = True
            End If
            If steps = 2 And eventhandled = False Then
                SendMessage("txtStatus", "Examining Files' Content")
                steps = 3
                txtFileQueue1.Text = 0

                txtFileQueueTotal1.Text = "(of " & filequeue1.Count & ")"


                thread1snapshot = 0


                If filequeue1.Count > 0 Then
                    Worker1.filequeue1 = CStr(filequeue1.Dequeue())
                    Try
                        Worker1.ChooseThreads(31)
                    Catch errt As Exception
                    End Try
                End If

                workerbusy = True
                eventhandled = True
            End If
            If steps = 3 And eventhandled = False Then
                Select Case queue
                    Case 1
                        If filequeue1.Count > 0 Then
                            Worker1.filequeue1 = CStr(filequeue1.Dequeue())
                            Try
                                Worker1.ChooseThreads(31)
                            Catch errt As Exception
                            End Try
                        End If

                End Select

                If filequeue1.Count = 0 Then
                    SendMessage("txtStatusBar1", "")
                    SendMessage("txtStatusBar2", "")
                    SendMessage("txtStatusBar3", "")
                    SendMessage("txtStatusBar4", "")
                    SendMessage("txtStatusBar5", "")
                    SendMessage("txtStatus", "Files' Content Examination Complete")
                    steps = 4
                    workerbusy = False

                Else
                    workerbusy = True
                End If
                eventhandled = True

            End If

            If steps = 4 And eventhandled = False Then
                SendMessage("txtStatusBar1", "")
                SendMessage("txtStatusBar2", "")
                SendMessage("txtStatusBar3", "")
                SendMessage("txtStatusBar4", "")
                SendMessage("txtStatusBar5", "")
                steps = 5
                Worker1.ChooseThreads(4)
                workerbusy = True
                eventhandled = True
            End If

            If steps = 5 And eventhandled = False Then
                WorkerStepAnnounceHandler(5)

                txtProcessEnded.Text = "Ended: " & Format(Now(), "dd/MM/yyyy HH:mm:ss")
                SendMessage("txtStatus", "Process Completed")
                workerbusy = False
                eventhandled = True
                txtSearchString.Select()
                ButtonOperationLaunch.Enabled = True
            End If

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub WorkerFileProcessingHandler(ByVal filename As String, ByVal queue As Integer)
        Try

            Select Case queue
                Case 1
                    SendMessage("txtStatusBar1", ("Processing: " & filename.Replace(txtBaseFolder.Text & "\", "...\")).ToUpper)
                    SendMessage("txtFileQueue1", CInt(txtFileQueue1.Text) + 1)

            End Select

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Shutting_Down()
        Try
            Save_Registry_Values()
            Worker1.Dispose()
            'splash_loader.Close()
            'Me.Close()
        Catch ex As Exception
            Error_Handler(ex)
            'Me.Close()
        End Try
    End Sub

    Private Sub Main_Screen_Close(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Shutting_Down()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        application_status_check()
    End Sub

    Private Sub statuslabel()
        Label14.Text = steps & " " & filequeue1.Count & " " & workerbusy.ToString
    End Sub

    Private Sub application_status_check()

        Try
            If steps = 3 Then
                If filequeue1.Count > 0 Then
                    If filequeue1.Count > 0 Then
                        If CLng(txtFileQueue1.Text) <= thread1snapshot Then
                            If Worker1.filequeue1 = "" Then
                                Worker1.filequeue1 = CStr(filequeue1.Dequeue())
                            End If
                            Try
                                SendMessage("txtFileQueue1", CLng(txtFileQueueTotal1.Text) - filequeue1.Count - 1)
                                Worker1.ChooseThreads(31)
                            Catch errt As Exception
                            End Try
                        End If
                        thread1snapshot = txtFileQueue1.Text
                    End If

                End If
            End If
            If steps = 4 And workerbusy = False Then
                WorkerCompleteHandler(0)
            End If
        Catch ex As Exception
            Error_Handler(ex, "Restart Stalled Threads")
        End Try
    End Sub


    Private Function load_datatypes() As Boolean
        Dim result As Boolean = False
        Try
            Dim path1, path2, filetoread As String
            path1 = (Application.StartupPath & "\config.ini").Replace("\\", "\")
            path2 = (Application.StartupPath & "\default_config.ini").Replace("\\", "\")
            filetoread = ""
            Dim finfo As FileInfo = New FileInfo(path1)
            If finfo.Exists = True Then
                filetoread = path1
            Else
                finfo = New FileInfo(path2)
                If finfo.Exists = True Then filetoread = path2
            End If

            If filetoread.Length > 0 Then
                filetypes = New ArrayList
                filetypes.Clear()

                Dim config As StreamReader = New StreamReader(filetoread)

                While config.Peek > -1
                    filetypes.Add(config.ReadLine.Trim)

                End While
                config.Close()
                filetypes.Sort()
                Dim str As String
                For Each str In filetypes
                    ListBox1.Items.Add(str)
                Next
                result = True
            End If


        Catch ex As Exception
            Error_Handler(ex)
            result = False
        End Try
        Return result
    End Function


    Private Sub Main_Screen_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Control.CheckForIllegalCrossThreadCalls = False
        Me.Text = "Find And Replace Single (Build " & My.Application.Info.Version.Major & Format(My.Application.Info.Version.Minor, "00") & Format(My.Application.Info.Version.Build, "00") & "." & Format(My.Application.Info.Version.Revision, "00") & ")"
        Load_Registry_Values()
        statuslabel()

        thread1snapshot = 0

        dataloaded = True
        'splash_loader.Visible = False
        Me.Show()
        If load_datatypes() = False Then
            Error_Handler("Error reading required Config.ini input file. This application will now shut down.")
            Shutting_Down()
        End If
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Worker1.WorkerThread Is Nothing = False Then
                Worker1.stopcheck = True
                WorkerStepAnnounceHandler(0)
                Worker1.WorkerThread.Abort()
                workerbusy = False
                SendMessage("txtStatus", "Process Terminated")
                ButtonOperationLaunch.Enabled = True
            End If
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Public Sub Load_Registry_Values()
        Try
            Dim configflag As Boolean
            configflag = False
            Dim str As String
            Dim keyflag1 As Boolean = False
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim keys() As String = oReg.GetSubKeyNames()
            System.Array.Sort(keys)

            For Each str In keys
                If str.Equals("Software\CodeUnit\Find And Replace") = True Then
                    keyflag1 = True
                    Exit For
                End If
            Next str

            If keyflag1 = False Then
                oReg.CreateSubKey("Software\CodeUnit\Find And Replace")
            End If

            keyflag1 = False

            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\CodeUnit\Find And Replace", True)

            str = oKey.GetValue("frmbasefolder")
            If Not IsNothing(str) And Not (str = "") Then
                frmbasefolder = str
            Else
                configflag = True
                oKey.SetValue("frmbasefolder", (Application.StartupPath))
                frmbasefolder = (Application.StartupPath)
            End If
            txtBaseFolder.Text = frmbasefolder


            str = oKey.GetValue("frmsearchstring")
            If Not IsNothing(str) And Not (str = "") Then
                frmsearchstring = str
            Else
                configflag = True
                oKey.SetValue("frmsearchstring", "String to Search for")
                frmsearchstring = ("String to Search for")
            End If
            txtSearchString.Text = frmsearchstring


            str = oKey.GetValue("frmreplacestring")
            If Not IsNothing(str) And Not (str = "") Then
                frmreplacestring = str
            Else
                configflag = True
                oKey.SetValue("frmreplacestring", "Replacement String")
                frmreplacestring = ("Replacement String")
            End If
            txtReplaceString.Text = frmreplacestring
            oKey.Close()
            oReg.Close()

        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub

    Private Sub Save_Registry_Values()
        Try
            Dim oReg As RegistryKey = Registry.LocalMachine
            Dim oKey As RegistryKey = oReg.OpenSubKey("Software\CodeUnit\Find And Replace", True)

            oKey.SetValue("frmsearchstring", txtSearchString.Text)
            oKey.SetValue("frmreplacestring", txtReplaceString.Text)
            oKey.SetValue("frmbasefolder", txtBaseFolder.Text)

            oKey.Close()
            oReg.Close()
        Catch ex As Exception
            Error_Handler(ex)
        End Try
    End Sub



    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        statuslabel()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        application_status_check()
    End Sub


End Class
