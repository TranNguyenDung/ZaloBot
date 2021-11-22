﻿Imports System.Diagnostics.Eventing
Imports System.IO
Imports System.IO.Ports
Imports System.Reflection




Public Class Form1
    'Set form
    Private Const VER = "1.5 "   'Add Find comp altium
    'Private Const VER = "1.4"   'Add Find comp altium
    'Private Const VER = "1.3"   'Add event serial recive data
    'Private Const VER = "1.2"   'Check data on windown
    'Private Const VER = "1.1"   'Begin add name com
    'Private Const VER = "1.0"   'Begin add name com

    Private Const FORM_WIDTH = 700
    Private Const FORM_HEIGHT = 550

    Private Const FIND_FOOTPRINT = "Find FootPrint"

    Private Coler1 As Color
    Private Coler2 As Color
    Private Coler3 As Color
    Private Coler4 As Color

    Private objPort As SerialPort
    Dim WithEvents serialPort As New SerialPort
    'Dim serialPort = New SerialPort

    'Delay
    Private WithEvents Timer1 As Windows.Forms.Timer
    Private WithEvents Timer2 As Windows.Forms.Timer

    'Menu
    Private mnuBar As MainMenu
    Private myMenuItemMenu As MenuItem
    Private myMenuItemOpen As MenuItem
    Private myMenuItemSave As MenuItem
    Private myMenuItemExit As MenuItem
    Private FolderBrowserDialog1 As FolderBrowserDialog
    Private OpenFileDialog1 As OpenFileDialog
    Private SaveFileDialog1 As SaveFileDialog

    'Com port
    Private cmbCOM As ComboBox
    Private lblComport As Label
    Private cmbBaud As ComboBox
    Private StrComName As String
    Private btnScanComPort As Button

    Private lblStatus As Label

    'Private with_name As Object
    Private grpPartition() As GroupBox

    Private lblTarget As Label
    Private txtTarget As TextBox
    Private cmbTarget As ComboBox

    Private btnEnaDisCOM As Button

    Private btnHide As Button


    Private btnAction() As Button
    Private cmbCtrl() As ComboBox
    Private cmbShift() As ComboBox
    Private chkCtrl() As CheckBox
    Private chkShift() As CheckBox
    Private cmbKey() As ComboBox

    Private lblSource As Label
    Private txtSource As TextBox
    Private lblTypeDo As Label
    Private cmbTypeDo As ComboBox
    Private lblSourceRef As Label
    Private txtSourceRef As TextBox
    Private btnDo As Button

    Private lblMouseLocation As Label
    Private txtMouseLocation() As TextBox
    Private lblSpace As Label
    Private txtSpace As TextBox
    Private txtSpaceHight As TextBox
    Private btnMouseLocation As Button

    Private btnScanApp As Button
    Private txtList As TextBox

    Private btnFindNext As Button
    Private btnFindBack As Button

    Private intFindFootprint As Integer = 0


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        interfaceDefaul()
        interfaceCustom()
        initData()
        init_gui()

    End Sub
    Private Sub interfaceDefaul()
        'MsgBox("hello")
        With Me

            Timer1 = New Timer()
            Timer2 = New Timer()

            mnuBar = New MainMenu()
            myMenuItemMenu = New MenuItem()
            myMenuItemOpen = New MenuItem()
            myMenuItemSave = New MenuItem()
            myMenuItemExit = New MenuItem()
            OpenFileDialog1 = New OpenFileDialog()
            SaveFileDialog1 = New SaveFileDialog()

            cmbCOM = New ComboBox()
            btnScanComPort = New Button()
            lblStatus = New Label()
            Coler1 = New Color
            Coler2 = New Color
            Coler3 = New Color
            Coler4 = New Color

            Coler1 = Color.AliceBlue
            Coler2 = Color.Chocolate
            Coler3 = Color.BlueViolet
            Coler4 = Color.Aqua


            btnInit(btnHide)

            'Group all
            ReDim grpPartition(0 To 2)
            For i = 0 To grpPartition.Count - 1
                grpPartition(i) = New GroupBox
            Next

            lblInit(lblTarget)
            txtInit(txtTarget)
            cmbInit(cmbTarget)

            btnInit(btnEnaDisCOM)
        End With

        With Me
            myMenuItemMenu = New MenuItem("Menu")
            myMenuItemOpen = New MenuItem("Open")
            myMenuItemSave = New MenuItem("Save")
            myMenuItemExit = New MenuItem("Exit")

            mnuBar.MenuItems.Add(myMenuItemMenu)
            myMenuItemMenu.MenuItems.Add(myMenuItemOpen)
            myMenuItemMenu.MenuItems.Add(myMenuItemSave)
            myMenuItemMenu.MenuItems.Add(myMenuItemExit)
            .Menu = mnuBar


            'Coler1 = SystemColors.ButtonFace
            'Coler2 = SystemColors.ActiveCaption
            'Coler3 = SystemColors.ActiveBorder
            Coler1 = SystemColors.ButtonFace
            Coler2 = Color.Chocolate
            Coler3 = Color.BlueViolet
            Coler4 = Color.Aqua

            .Text = "SendKey_Ver " & VER
            .ControlBox = True
            .MaximizeBox = False
            .MinimizeBox = True
            .FormBorderStyle = FormBorderStyle.FixedSingle
            .Name = "SendKey"
            '.Size = New Size(location_and_size_tab_DACS2_x + 40, location_and_size_tab_DACS2_y + 250)
            .Size = New Size(FORM_WIDTH, FORM_HEIGHT)
            .StartPosition = FormStartPosition.CenterScreen 'center screen
            .BackColor = New Color
            REM:*************************************
            Dim lc_x As Integer
            Dim lc_y As Integer
            Dim Size_x As Integer
            Dim size_y As Integer

            lc_x = 20
            lc_y = 10
            Size_x = 150
            size_y = 30
            lbl_set(lblStatus, "lblStatus", "Status", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(lblStatus)
            REM:*************************************
            lc_x = lblStatus.Location.X + lblStatus.Width + 1
            lc_y = lblStatus.Location.Y
            Size_x = 90
            size_y = lblStatus.Size.Height
            cmb_set(cmbCOM, "cmbCOM", "", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(cmbCOM)
            REM:*************************************
            lc_x = cmbCOM.Location.X + cmbCOM.Width + 1
            lc_y = cmbCOM.Location.Y
            Size_x = 80
            size_y = lblStatus.Size.Height
            btn_set(btnScanComPort, "btnScanComPort", "Refresh", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(btnScanComPort)

            'btnHide
            REM:*************************************
            lc_x = btnScanComPort.Location.X + btnScanComPort.Width + 1
            lc_y = btnScanComPort.Location.Y
            Size_x = 55
            size_y = lblStatus.Size.Height
            btn_set(btnHide, "btnHide", "Hide", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(btnHide)

            'lblInit(lblTarget)

            REM:*************************************
            lc_x = btnHide.Location.X + btnHide.Width + 1
            lc_y = btnHide.Location.Y
            Size_x = 120
            size_y = lblStatus.Size.Height
            'btnInit(btnEnaDisCOM)
            btn_set(btnEnaDisCOM, "btnEnaDisCOM", "Connect COM", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(btnEnaDisCOM)
            'lbl_set(lblTarget, "lblTarget", "Target:->", 12, Coler1, Size_x, size_y, lc_x, lc_y)
            '.Controls.Add(lblTarget)

            'txtInit(txtTarget)
            lc_x = btnEnaDisCOM.Location.X + btnEnaDisCOM.Width + 1
            lc_y = btnEnaDisCOM.Location.Y
            Size_x = 155
            size_y = lblStatus.Size.Height
            'txt_set(txtTarget, "txtTarget", "-----------------------", 12, Size_x, size_y, lc_x, lc_y)
            cmb_set(cmbTarget, "cmbTarget", "", 12, Coler4, Size_x, size_y, lc_x, lc_y)

            .Controls.Add(cmbTarget)

            REM:*************************************
            lc_x = 20
            lc_y = 41
            Size_x = 440
            size_y = 460
            grp_set(grpPartition(0), "grpPartition00", "", 10.7, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(grpPartition(0))
            REM:*************************************
            lc_x = grpPartition(0).Location.X + grpPartition(0).Width + 5
            lc_y = grpPartition(0).Location.Y ' + grpPartition(0).Height + 7
            Size_x = 210
            size_y = 460
            grp_set(grpPartition(1), "grpPartition01", "", 10.7, Size_x, size_y, lc_x, lc_y)
            .Controls.Add(grpPartition(1))

            REM:************************************************************************************

        End With


        '-----------------------------------------------------------------------------------------------------------------------------
        AddHandler btnScanComPort.Click, AddressOf btn_click
        AddHandler cmbCOM.SelectedIndexChanged, AddressOf cmbCOM_Changed
        AddHandler btnHide.Click, AddressOf btn_click
        AddHandler btnEnaDisCOM.Click, AddressOf btn_click
        AddHandler myMenuItemOpen.Click, AddressOf MenuItemHandler
        AddHandler myMenuItemSave.Click, AddressOf MenuItemHandler
        AddHandler myMenuItemExit.Click, AddressOf MenuItemHandler
    End Sub

    Private Sub MenuItemHandler(sender As Object, e As EventArgs)

        Dim MenuItemClick As MenuItem = TryCast(sender, MenuItem)

        If MenuItemClick IsNot Nothing Then
            Dim strText As String = MenuItemClick.Text
            Dim strDmy As String = ""
            If strText.StartsWith("Open") Then
                Debug.WriteLine(MenuItemClick.Text)
                OpenFileDialog1.DefaultExt = "TriggerKey"
                OpenFileDialog1.Filter = ".TriggerKey files (*.TriggerKey)|*.TriggerKey" & "|All Files|*.*"
                If (OpenFileDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim strName As String = Path.GetFileName(OpenFileDialog1.FileName)
                    Dim strPath As String = Path.GetDirectoryName(OpenFileDialog1.FileName)
                    'Debug.WriteLine(strPath)
                    ReadFile(strPath, strName)
                End If
            ElseIf strText.StartsWith("Save") Then
                SaveFileDialog1.Filter = "TriggerKey|*.TriggerKey"
                SaveFileDialog1.Title = "Save an TriggerKey File"
                If (SaveFileDialog1.ShowDialog() = DialogResult.OK) Then
                    Dim strName As String = Path.GetFileName(SaveFileDialog1.FileName)
                    Dim strPath As String = Path.GetDirectoryName(SaveFileDialog1.FileName)
                    'Debug.WriteLine(MenuItemClick.Text)
                    'Debug.WriteLine(strName)
                    'Debug.WriteLine(strPath)
                    WriteFile(strPath, strName)
                End If
            ElseIf strText.StartsWith("Exit") Then
                Me.Close()

                Debug.WriteLine(MenuItemClick.Text)
            End If
        End If
    End Sub

    Private Sub interfaceCustom()
        With Me
            btnInit(btnAction, 16)
            cmbInit(cmbCtrl, 16)
            cmbInit(cmbShift, 16)
            chkInit(chkCtrl, 16)
            chkInit(chkShift, 16)
            cmbInit(cmbKey, 16)

            btnInit(btnScanApp)
            txtInit(txtList)

            lblInit(lblSource)
            lblInit(lblTypeDo)
            txtInit(txtSource)
            txtInit(txtSourceRef)
            cmbInit(cmbTypeDo)
            btnInit(btnDo)
            lblInit(lblMouseLocation)
            txtInit(txtMouseLocation, 2)
            lblInit(lblSpace)
            txtInit(txtSpace)
            txtInit(txtSpaceHight)
            btnInit(btnMouseLocation)
            btnInit(btnFindNext)
            btnInit(btnFindBack)

        End With

        With Me
            Dim lc_x As Integer
            Dim lc_y As Integer
            Dim s_x As Integer
            Dim s_y As Integer

            For i = 0 To btnAction.Count - 1
                s_x = 150
                s_y = 25
                lc_x = 10
                lc_y = 20 + (s_y + 2) * i
                btn_set(btnAction(i), "btnAction" & Format(i, "00"), "Action Key No: " & i.ToString, 10, Coler1, s_x, s_y, lc_x, lc_y)
                grpPartition(0).Controls.Add(btnAction(i))
                btnAction(i).Visible = False
            Next

            'chkInit(chkCtrl, 16)

            For i = 0 To chkCtrl.Count - 1
                s_x = 50
                s_y = 25
                lc_x = btnAction(i).Location.X + btnAction(i).Width + 5
                lc_y = btnAction(i).Location.Y
                chk_set(chkCtrl(i), "chkCtrl", "Ctrl", 10, False, s_x, s_y, lc_x, lc_y)
                grpPartition(0).Controls.Add(chkCtrl(i))
                chkCtrl(i).Visible = False
            Next

            'chkInit(chkShift, 16)
            For i = 0 To chkShift.Count - 1
                s_x = 55
                s_y = 25
                lc_x = chkCtrl(i).Location.X + chkCtrl(i).Width + 5
                lc_y = chkCtrl(i).Location.Y
                chk_set(chkShift(i), "chkShift", "Shift", 10, False, s_x, s_y, lc_x, lc_y)
                grpPartition(0).Controls.Add(chkShift(i))
                chkShift(i).Visible = False
            Next

            'cmbInit(cmbKey, 16)
            For i = 0 To cmbKey.Count - 1
                s_x = 150
                s_y = 25
                lc_x = chkShift(i).Location.X + chkShift(i).Width + 5
                lc_y = chkShift(i).Location.Y
                cmb_set(cmbKey(i), "chkShift", "Shift", 10, Color.AliceBlue, s_x, s_y, lc_x, lc_y)
                grpPartition(0).Controls.Add(cmbKey(i))
                cmbKey(i).Visible = False
            Next


            'Private lblSource As Label
            'Private txtSource As TextBox
            'Private lblTypeDo As Label
            'Private cmbTypeDo As ComboBox
            'Private lblSourceRef As Label
            'Private txtSourceRef As TextBox
            s_x = 150
            s_y = 25
            lc_x = 10
            lc_y = 20
            lbl_set(lblSource, "lblSource", "Source", 12, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(lblSource)
            s_x = 150
            s_y = 400
            lc_x = lblSource.Location.X
            lc_y = lblSource.Location.Y + lblSource.Height + 5
            txt_set(txtSource, "txtSource", "", 10, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(txtSource)
            txtSource.Multiline = True

            s_x = 200
            s_y = 25
            lc_x = lblSource.Location.X + lblSource.Width + 5
            lc_y = lblSource.Location.Y
            lbl_set(lblTypeDo, "lblTypeDo", "Type Do", 12, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(lblTypeDo)

            s_x = lblTypeDo.Width
            s_y = 25
            lc_x = lblTypeDo.Location.X ' + lblTypeDo.Width + 5
            lc_y = lblTypeDo.Location.Y + lblTypeDo.Height + 5
            cmb_set(cmbTypeDo, "cmbTypeDo", "", 12, Coler4, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(cmbTypeDo)
            cmbTypeDo.Items.Clear()
            cmbTypeDo.Items.Add("D2.Change FootPrint")
            cmbTypeDo.Items.Add("D2.Change Name")
            cmbTypeDo.Items.Add("D2.Hidden Value")
            cmbTypeDo.Items.Add("Al.MoveComp")
            cmbTypeDo.Items.Add("RF007.TX-Idle")
            cmbTypeDo.Items.Add(FIND_FOOTPRINT)
            cmbTypeDo.SelectedIndex = 0
            AddHandler cmbTypeDo.SelectedIndexChanged, AddressOf cmbTypeDo_SelectIndexChange

            s_x = cmbTypeDo.Width
            s_y = 25
            lc_x = cmbTypeDo.Location.X
            lc_y = cmbTypeDo.Location.Y + cmbTypeDo.Height + 5
            txt_set(txtSourceRef, "txtSourceRef", "", 10, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(txtSourceRef)

            'btnDo
            s_x = 70
            s_y = 90
            lc_x = lblTypeDo.Location.X + lblTypeDo.Width + 5
            lc_y = lblTypeDo.Location.Y
            btn_set(btnDo, "btnDo", "Apply", 10, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(btnDo)
            AddHandler btnDo.Click, AddressOf btn_click

            'lblInit(lblMouseLocation)
            s_x = 90
            s_y = 25
            lc_x = txtSourceRef.Location.X ' + txtSourceRef.Width + 5
            lc_y = txtSourceRef.Location.Y + txtSourceRef.Height + 5
            lbl_set(lblMouseLocation, "lblMouseLocation", "MouseXY", 12, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(lblMouseLocation)

            For i = 0 To txtMouseLocation.Count - 1
                s_x = 90
                s_y = 25
                lc_x = lblMouseLocation.Location.X + lblMouseLocation.Width + 2 + i * s_x
                lc_y = lblMouseLocation.Location.Y ' + lblMouseLocation.Height + 5
                txt_set(txtMouseLocation(i), "txtMouseLocation" & i.ToString, "200", 10, s_x, s_y, lc_x, lc_y)
                grpPartition(0).Controls.Add(txtMouseLocation(i))
            Next

            'lblInit(lblSpace)
            s_x = 90
            s_y = 25
            lc_x = lblMouseLocation.Location.X '+ lblMouseLocation.Width + 5
            lc_y = lblMouseLocation.Location.Y + lblMouseLocation.Height + 5
            lbl_set(lblSpace, "lblSpace", "SpaceXY", 12, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(lblSpace)

            'txtInit(txtSpace)
            s_x = 90
            s_y = 25
            lc_x = lblSpace.Location.X + lblSpace.Width + 2
            lc_y = lblSpace.Location.Y ' + lblSpace.Height + 5
            txt_set(txtSpace, "txtSpace", "20", 10, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(txtSpace)
            'txtSpaceHight
            s_x = 90
            s_y = 25
            lc_x = txtSpace.Location.X + txtSpace.Width
            lc_y = txtSpace.Location.Y ' + txtSpace.Height + 5
            txt_set(txtSpaceHight, "txtSpaceHight", "30", 10, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(txtSpaceHight)

            'btnInit(btnMouseLocation)
            s_x = 200
            s_y = 50
            lc_x = lblSpace.Location.X '+ lblSpace.Width + 2
            lc_y = lblSpace.Location.Y + lblSpace.Height + 5
            btn_set(btnMouseLocation, "btnMouseLocation", "Get location Mouse 5s", 10, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(btnMouseLocation)
            AddHandler btnMouseLocation.Click, AddressOf btn_click

            'btnInit(btnFindNext)
            s_x = 130
            s_y = 50
            lc_x = btnMouseLocation.Location.X '+ btnMouseLocation.Width + 2
            lc_y = btnMouseLocation.Location.Y + btnMouseLocation.Height + 5
            btn_set(btnFindNext, "btnFindNext", "FindNext", 10, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(btnFindNext)
            AddHandler btnFindNext.Click, AddressOf btn_click

            'btnInit(btnFindBack)
            s_x = 130
            s_y = 50
            lc_x = btnFindNext.Location.X + btnFindNext.Width + 2
            lc_y = btnFindNext.Location.Y ' + btnFindNext.Height + 5
            btn_set(btnFindBack, "btnFindBack", "FindBack", 10, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(0).Controls.Add(btnFindBack)
            AddHandler btnFindBack.Click, AddressOf btn_click

            'btnInit(btnScanApp)
            s_x = grpPartition(1).Width - 10
            s_y = 25
            lc_x = 5
            lc_y = 20
            btn_set(btnScanApp, "btnScanApp", "Scan App", 10, Coler1, s_x, s_y, lc_x, lc_y)
            grpPartition(1).Controls.Add(btnScanApp)


            'txtInit(txtList)
            s_x = grpPartition(1).Width - 10
            s_y = 400
            lc_x = 5
            lc_y = btnScanApp.Location.Y + btnScanApp.Height + 5
            txt_set(txtList, "txtList", "", 10, s_x, s_y, lc_x, lc_y)
            grpPartition(1).Controls.Add(txtList)
            txtList.Multiline = True
        End With
        For i = 0 To btnAction.Count - 1
            AddHandler btnAction(i).Click, AddressOf btn_click
        Next
        AddHandler btnScanApp.Click, AddressOf btn_click
    End Sub

    Private Sub initData()



    End Sub

    Private Sub init_gui()
        Dim i As Integer = 0
        '----------------------
        Key.init()

        For i = 0 To cmbKey.Count - 1
            cmbKey(i).Items.Clear()


            cmbKey(i).Items.Add(Key.A)
            cmbKey(i).Items.Add(Key.B)
            cmbKey(i).Items.Add(Key.C)
            cmbKey(i).Items.Add(Key.D)
            cmbKey(i).Items.Add(Key.E)
            cmbKey(i).Items.Add(Key.F)
            cmbKey(i).Items.Add(Key.G)
            cmbKey(i).Items.Add(Key.H)
            cmbKey(i).Items.Add(Key.I)
            cmbKey(i).Items.Add(Key.J)
            cmbKey(i).Items.Add(Key.K)
            cmbKey(i).Items.Add(Key.L)
            cmbKey(i).Items.Add(Key.M)
            cmbKey(i).Items.Add(Key.N)
            cmbKey(i).Items.Add(Key.O)
            cmbKey(i).Items.Add(Key.P)
            cmbKey(i).Items.Add(Key.Q)
            cmbKey(i).Items.Add(Key.R)
            cmbKey(i).Items.Add(Key.S)
            cmbKey(i).Items.Add(Key.T)
            cmbKey(i).Items.Add(Key.U)
            cmbKey(i).Items.Add(Key.V)
            cmbKey(i).Items.Add(Key.W)
            cmbKey(i).Items.Add(Key.X)
            cmbKey(i).Items.Add(Key.Y)
            cmbKey(i).Items.Add(Key.Z)

            cmbKey(i).Items.Add(Key.N0)
            cmbKey(i).Items.Add(Key.N1)
            cmbKey(i).Items.Add(Key.N2)
            cmbKey(i).Items.Add(Key.N3)
            cmbKey(i).Items.Add(Key.N4)
            cmbKey(i).Items.Add(Key.N5)
            cmbKey(i).Items.Add(Key.N6)
            cmbKey(i).Items.Add(Key.N7)
            cmbKey(i).Items.Add(Key.N8)
            cmbKey(i).Items.Add(Key.N9)

            cmbKey(i).Items.Add(Key.F1)
            cmbKey(i).Items.Add(Key.F2)
            cmbKey(i).Items.Add(Key.F3)
            cmbKey(i).Items.Add(Key.F4)
            cmbKey(i).Items.Add(Key.F5)
            cmbKey(i).Items.Add(Key.F6)
            cmbKey(i).Items.Add(Key.F7)
            cmbKey(i).Items.Add(Key.F8)
            cmbKey(i).Items.Add(Key.F9)
            cmbKey(i).Items.Add(Key.F10)
            cmbKey(i).Items.Add(Key.F11)
            cmbKey(i).Items.Add(Key.F12)
            cmbKey(i).Items.Add(Key.F13)
            cmbKey(i).Items.Add(Key.F14)
            cmbKey(i).Items.Add(Key.F15)
            cmbKey(i).Items.Add(Key.F16)
            cmbKey(i).Items.Add(Key.F17)
            cmbKey(i).Items.Add(Key.F18)
            cmbKey(i).Items.Add(Key.F19)
            cmbKey(i).Items.Add(Key.F20)
            cmbKey(i).Items.Add(Key.F21)
            cmbKey(i).Items.Add(Key.F22)
            cmbKey(i).Items.Add(Key.F23)
            cmbKey(i).Items.Add(Key.F24)

            cmbKey(i).Items.Add(Key.ADD)
            cmbKey(i).Items.Add(Key.SUBTRACT)
            cmbKey(i).Items.Add(Key.MULTIPLY)
            cmbKey(i).Items.Add(Key.DIVIDE)
            cmbKey(i).Items.Add(Key.BACKSPACE)
            cmbKey(i).Items.Add(Key.BREAK)
            cmbKey(i).Items.Add(Key.APSLOCK)
            cmbKey(i).Items.Add(Key.DELETE)
            cmbKey(i).Items.Add(Key.UP)
            cmbKey(i).Items.Add(Key.DOWN)
            cmbKey(i).Items.Add(Key.KEND)
            cmbKey(i).Items.Add(Key.ENTER)
            cmbKey(i).Items.Add(Key.ESC)
            cmbKey(i).Items.Add(Key.HELP)
            cmbKey(i).Items.Add(Key.HOME)
            cmbKey(i).Items.Add(Key.INSERT)
            cmbKey(i).Items.Add(Key.LEFT)
            cmbKey(i).Items.Add(Key.NUMLOCK)
            cmbKey(i).Items.Add(Key.PAGEDOWN)
            cmbKey(i).Items.Add(Key.PAGEUP)
            cmbKey(i).Items.Add(Key.PRINTSCREEN)
            cmbKey(i).Items.Add(Key.RIGHT)
            cmbKey(i).Items.Add(Key.SCROLLLOCK)
            cmbKey(i).Items.Add(Key.TAB)
            cmbKey(i).Items.Add(Key.SHIFT)
            cmbKey(i).Items.Add(Key.ALT)
        Next
        For i = 0 To cmbKey.Count - 1
            cmbKey(i).SelectedIndex = i
        Next

        cmbTarget.Items.Add("App Activing")
        cmbTarget.SelectedIndex = 0
    End Sub

    Private Sub grpInit(ByRef grpInit As GroupBox)
        grpInit = New GroupBox()
    End Sub

    Private Sub rdobtn(ByRef rdobtn As RadioButton)
        rdobtn = New RadioButton()
    End Sub

    Private Sub rdobtn(ByRef rdobtn() As RadioButton, ByVal len As Integer)
        ReDim rdobtn(0 To len - 1)
        For i = 0 To len - 1
            rdobtn(i) = New RadioButton()
        Next
    End Sub

    Private Sub lblInit(ByRef lbl As Label)
        lbl = New Label()
    End Sub

    Private Sub lblInit(ByRef lbl() As Label, ByVal len As Integer)
        ReDim lbl(0 To len - 1)
        For i = 0 To len - 1
            lbl(i) = New Label()
        Next
    End Sub

    Private Sub cmbInit(ByRef cmb As ComboBox)
        cmb = New ComboBox()
    End Sub

    Private Sub cmbInit(ByRef cmb() As ComboBox, ByVal len As Integer)
        ReDim cmb(0 To len - 1)
        For i = 0 To len - 1
            cmb(i) = New ComboBox()
        Next
    End Sub

    Private Sub chkInit(ByRef chk() As CheckBox, ByVal len As Integer)
        ReDim chk(0 To len - 1)
        For i = 0 To len - 1
            chk(i) = New CheckBox()
        Next
    End Sub

    Private Sub btnInit(ByRef btn As Button)
        btn = New Button()
    End Sub

    Private Sub btnInit(ByRef btn() As Button, len As Integer)
        ReDim btn(0 To len - 1)
        For i = 0 To len - 1
            btn(i) = New Button()
        Next
    End Sub

    Private Sub txtInit(ByRef txt As TextBox)
        txt = New TextBox()
    End Sub

    Private Sub txtInit(ByRef txt() As TextBox, len As Integer)
        ReDim txt(0 To len - 1)
        For i = 0 To len - 1
            txt(i) = New TextBox()
        Next
    End Sub

    Private Sub lbl_set(with_name As Object, Name As String, Text As String, text_size As Single, cColor As Color, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(127, Byte))
            .BackColor = cColor
            .TabStop = False
            .BorderStyle = BorderStyle.FixedSingle
            '.TextAlign = ContentAlignment.MiddleCenter
            .TextAlign = ContentAlignment.MiddleLeft
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
        End With
    End Sub

    Private Sub cmb_set(with_name As Object, Name As String, Text As String, text_size As Single, cColor As Color, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(127, Byte))
            .DropDownStyle = ComboBoxStyle.DropDownList
            .BackColor = Coler1
            .TabStop = False
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
        End With
    End Sub

    Private Sub btn_set(with_name As Object, Name As String, Text As String, text_size As Single, cColor As Color, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(127, Byte))
            .BackColor = Coler1
            .TabStop = False
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
        End With
    End Sub

    Private Sub grp_set(with_name As Object, Name As String, Text As String, text_size As Single, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
            .BackgroundImageLayout = ImageLayout.Center
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(50, Byte))
        End With
    End Sub

    Private Sub chk_set(with_name As Object, Name As String, Text As String, text_size As Single, check As Boolean, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
            '.BackgroundImageLayout = ImageLayout.Center
            '.TextAlign = ContentAlignment.MiddleCenter
            .TextAlign = ContentAlignment.MiddleLeft
            '.TabIndex = j
            .Checked = check
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(50, Byte))
        End With
    End Sub

    Private Sub txt_set(with_name As TextBox, Name As String, Text As String, text_size As Single, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
            '.BackgroundImageLayout = ImageLayout.Center
            '.TextAlign = ContentAlignment.MiddleCenter
            '.TextAlign = ContentAlignment.MiddleLeft
            '.TabIndex = j
            .Multiline = True
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(50, Byte))

        End With
    End Sub

    Private Sub rdobtn_set(with_name As RadioButton, Name As String, Text As String, text_size As Single, check As Boolean, sz_x As Integer, sz_y As Integer, lc_x As Integer, lc_y As Integer)
        With with_name
            .Name = Name
            .Text = Text
            .Size = New Size(sz_x, sz_y)
            .Location = New Point(lc_x, lc_y)
            '.BackgroundImageLayout = ImageLayout.Center
            '.TextAlign = ContentAlignment.MiddleCenter
            .TextAlign = ContentAlignment.MiddleLeft
            '.TabIndex = j
            .Checked = check
            .Font = New Font("ＭＳ", text_size, FontStyle.Regular, GraphicsUnit.Point, CType(50, Byte))
        End With
    End Sub



    Function Bit2Byte(b7 As Byte, b6 As Byte, b5 As Byte, b4 As Byte, b3 As Byte, b2 As Byte, b1 As Byte, b0 As Byte) As Byte
        Bit2Byte = 0
        If b7 Then
            Bit2Byte = Bit2Byte Or &H80
        End If
        If b6 Then
            Bit2Byte = Bit2Byte Or &H40
        End If
        If b5 Then
            Bit2Byte = Bit2Byte Or &H20
        End If
        If b4 Then
            Bit2Byte = Bit2Byte Or &H10
        End If
        If b3 Then
            Bit2Byte = Bit2Byte Or &H8
        End If
        If b2 Then
            Bit2Byte = Bit2Byte Or &H4
        End If
        If b1 Then
            Bit2Byte = Bit2Byte Or &H2
        End If
        If b0 Then
            Bit2Byte = Bit2Byte Or &H1
        End If
    End Function

    Function Dec2Hex(ByVal data As Integer) As String
        Dec2Hex = Hex(data)
        While Dec2Hex.Length < 2
            Dec2Hex = "0" & Dec2Hex
        End While
    End Function

    Private Function DEC2HEXString1Byte(intNumber As Byte) As String
        Dim number As Byte
        DEC2HEXString1Byte = "--"

        intNumber = intNumber And &HFF
        number = intNumber >> 4
        number = number And &HF
        DEC2HEXString1Byte = DEC2String4Bit(number)

        number = intNumber And &HF
        DEC2HEXString1Byte &= DEC2String4Bit(number)
    End Function

    Public Shared Function DEC2String4Bit(intNumber As Byte) As String
        Dim number As Byte

        number = intNumber And &HF

        If number < 10 Then
            DEC2String4Bit = number.ToString
        ElseIf number = 10 Then
            DEC2String4Bit = "A"
        ElseIf number = 11 Then
            DEC2String4Bit = "B"
        ElseIf number = 12 Then
            DEC2String4Bit = "C"
        ElseIf number = 13 Then
            DEC2String4Bit = "D"
        ElseIf number = 14 Then
            DEC2String4Bit = "E"
        ElseIf number = 15 Then
            DEC2String4Bit = "F"
        Else
            DEC2String4Bit = "-"
        End If

    End Function



    'Processing button
    Private Sub Read_Write(sender As Object, e As EventArgs)
        Dim strTx As String = ""
        Dim strRx As String = ""
        DisClick()
        Dim blStatus As Boolean = False
        setLabelText(lblStatus, "Processing...")

        Dim cButton As Button = TryCast(sender, Button)
        '-----------------------------------------------------------------
        If cButton IsNot Nothing Then
            Dim strDmy As String = cButton.Name
            If strDmy.StartsWith("btnWrite") Then
                blStatus = Write()
            ElseIf strDmy.StartsWith("bntRead") Then
                blStatus = Read()
            ElseIf strDmy.StartsWith("bntTXMode") Then
                'Debug.WriteLine("bntTXMode")


            ElseIf strDmy.StartsWith("bntRXMode") Then
                'Debug.WriteLine("bntRXMode")

            ElseIf strDmy.StartsWith("bntFreeMode") Then
                'Debug.WriteLine("bntFreeMode")

            ElseIf strDmy.StartsWith("bntTXBuff") Then
                Debug.WriteLine("bntTXBuff")

            ElseIf strDmy.StartsWith("bntRXBuff") Then
                Debug.WriteLine("bntRXBuff")

            ElseIf strDmy.StartsWith("btnCEHigh") Then
                Debug.WriteLine("btnCEHigh")
                blStatus = CEHIGH()
            ElseIf strDmy.StartsWith("btnCELow") Then
                Debug.WriteLine("btnCELow")
                blStatus = CELOW()
            ElseIf strDmy.StartsWith("btnFlushTX") Then
                Debug.WriteLine("btnFlushTX")
                blStatus = FlushTX()
            ElseIf strDmy.StartsWith("btnFlushRX") Then
                Debug.WriteLine("btnFlushRX")
                blStatus = FlushRX()

            End If
        End If
        '-----------------------------------------------------------------
        If blStatus = False Then
            setLabelText(lblStatus, "Error")
        Else
            setLabelText(lblStatus, "Processing OK")
        End If
        EnaClick()

    End Sub





    Function FlushTX() As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        FlushTX = True

        strTx = "FI~FLUSHTX" & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            FlushTX = False
        End If

    End Function

    Function FlushRX() As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        FlushRX = True

        strTx = "FI~FLUSHTX" & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            FlushRX = False
        End If

    End Function

    Function CEHIGH() As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        CEHIGH = True

        strTx = "FI~CEHIGH" & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            CEHIGH = False
        End If

    End Function

    Function CELOW() As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        CELOW = True

        strTx = "FI~CELOW" & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            CELOW = False
        End If

    End Function

    Function Write() As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        Dim b7 As Boolean = False
        Dim b6 As Boolean = False
        Dim b5 As Boolean = False
        Dim b4 As Boolean = False
        Dim b3 As Boolean = False
        Dim b2 As Boolean = False
        Dim b1 As Boolean = False
        Dim b0 As Boolean = False



        Write = True
        Exit Function
ERROR_WRITE:
        Write = False
        Exit Function
    End Function

    Function Read() As Boolean
        Dim i As Integer = 0
        Read = False
        Dim b7 As Integer = 0
        Dim b6 As Integer = 0
        Dim b5 As Integer = 0
        Dim b4 As Integer = 0
        Dim b3 As Integer = 0
        Dim b2 As Integer = 0
        Dim b1 As Integer = 0
        Dim b0 As Integer = 0

        Read = True
        Exit Function
ERROR_READ:
        Read = False
        Exit Function
    End Function

    Function Write_reg(Reg As Integer, byte4 As Byte, byte3 As Byte, byte2 As Byte, byte1 As Byte, byte0 As Byte) As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        Write_reg = True

        strTx = "FI~NRFW" & Dec2Hex(Reg) & Dec2Hex(byte0) & Dec2Hex(byte1) & Dec2Hex(byte2) & Dec2Hex(byte3) & Dec2Hex(byte4) & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            Write_reg = False
        End If

    End Function

    Function Write_reg(Reg As Integer, data As Byte) As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        Write_reg = True
        strTx = "FI~NRFW" & Dec2Hex(Reg) & Dec2Hex(data) & vbCrLf
        strRx = send_data(0, strTx)
        If strRx.StartsWith("OK") = False Then
            Write_reg = False
        End If

    End Function

    Function Read_reg(reg As Integer, ByRef data As Byte) As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        strTx = "FI~NRFR" & Dec2Hex(reg) & vbCrLf
        strRx = send_data(0, strTx)
        If strRx = "" Then
            GoTo ERROR_DATA
        End If
        data = CInt("&H" & strRx.Substring(0, 2))
        Read_reg = True
        Exit Function
ERROR_DATA:
        Read_reg = False
        Exit Function
    End Function

    Function Read_reg(reg As Integer, ByRef data() As Byte) As Boolean
        Dim strTx As String = ""
        Dim strRx As String = ""
        strTx = "FI~NRFR" & Dec2Hex(reg) & vbCrLf
        strRx = send_data(0, strTx)
        If strRx = "" Then
            GoTo ERROR_DATA
        End If
        data(0) = CInt("&H" & strRx.Substring(0, 2))
        data(1) = CInt("&H" & strRx.Substring(2, 2))
        data(2) = CInt("&H" & strRx.Substring(4, 2))
        data(3) = CInt("&H" & strRx.Substring(6, 2))
        data(4) = CInt("&H" & strRx.Substring(8, 2))
        Read_reg = True
        Exit Function
ERROR_DATA:
        Read_reg = False
        Exit Function
    End Function

    Private Sub DisClick()
        grpPartition(0).Enabled = False
        cmbCOM.Enabled = False
        btnScanComPort.Enabled = False
    End Sub

    Private Sub EnaClick()
        grpPartition(0).Enabled = True
        cmbCOM.Enabled = True
        btnScanComPort.Enabled = True
    End Sub


    Private Function CBit2Byte(b7 As Boolean, b6 As Boolean, b5 As Boolean, b4 As Boolean, b3 As Boolean, b2 As Boolean, b1 As Boolean, b0 As Boolean) As Byte
        CBit2Byte = &H0
        If b7 = True Then
            CBit2Byte = CBit2Byte Or &H80
        End If
        If b6 = True Then
            CBit2Byte = CBit2Byte Or &H40
        End If
        If b5 = True Then
            CBit2Byte = CBit2Byte Or &H20
        End If
        If b4 = True Then
            CBit2Byte = CBit2Byte Or &H10
        End If
        If b3 = True Then
            CBit2Byte = CBit2Byte Or &H8
        End If
        If b2 = True Then
            CBit2Byte = CBit2Byte Or &H4
        End If
        If b1 = True Then
            CBit2Byte = CBit2Byte Or &H2
        End If
        If b0 = True Then
            CBit2Byte = CBit2Byte Or &H1
        End If
    End Function

    Function Conv2Hex(Data As Integer) As String
        If Len(Hex(Data)) < 2 Then
            Conv2Hex = "0" & Hex(Data)
        Else
            Conv2Hex = Hex(Data)
        End If
    End Function

    Function Dec2Hex(Data As Integer, len As Integer) As String
        Dec2Hex = Hex(Data)
        While Dec2Hex.Length < len - 1
            Dec2Hex = "0" & Dec2Hex
        End While
    End Function


    Private Function send_data(no As Integer, strdata As String) As String
        'MsgBox(StrComName)
        send_data = send_data_serial(StrComName, 115200, 0, 8, 1, strdata)
    End Function

    Private Function send_data_serial(Name_ComPort As String, BaudRate As Integer, Parity As Integer, DataBits As Integer, StopBits As Integer, data_send As String) As String
        Dim strRx As String
        'Dim intLen As Integer

        strRx = ""


        ' Allow the user to set the appropriate properties.
        Dim serialPort = New SerialPort With {
            .PortName = Name_ComPort,
            .BaudRate = BaudRate,
            .Parity = Parity,
            .DataBits = DataBits,
            .StopBits = StopBits
        }

        If serialPort.IsOpen = True Then
            GoTo ERROR_END
        End If

        serialPort.Open()

        If serialPort.IsOpen = False Then
            GoTo ERROR_END
        End If

        Try
            'MsgBox(serialPort.IsOpen)

            serialPort.WriteTimeout = 100
            serialPort.ReadTimeout = 500
            serialPort.Write(data_send)

            Do
                strRx = serialPort.ReadLine
                If strRx Is Nothing Then
                    Exit Do
                ElseIf strRx.EndsWith(vbCr) Then
                    Exit Do
                End If
            Loop
        Catch ex As TimeoutException
            strRx = "Error: Serial Port read timed out."
        Finally
            If serialPort IsNot Nothing Then serialPort.Close()
        End Try

        send_data_serial = strRx
        serialPort.Close()
        Exit Function
ERROR_END:
        send_data_serial = ""
        Exit Function
        serialPort.Close()
    End Function

    Private Function Check_Status_Serial(Name_ComPort As String)
        Check_Status_Serial = True

        'objPort.PortName = Name_ComPort
        On Error GoTo ERROR_END
        If objPort.IsOpen = False Then
            Check_Status_Serial = False
        End If



        Check_Status_Serial = True
ERROR_END:

    End Function

    'Get name Com port
    Private Sub cmbCOM_Changed(sender As Object, e As EventArgs)
        Dim cmbComSelect As ComboBox = TryCast(sender, ComboBox)
        If cmbComSelect IsNot Nothing Then
            StrComName = cmbComSelect.Text
            'MsgBox(StrComName)
        End If
    End Sub


    Private Sub ScanComPort()
        Dim strDmy As String = ""
        cmbCOM.Items.Clear()
        For Each strDmy In My.Computer.Ports.SerialPortNames
            cmbCOM.Items.Add(strDmy)
            cmbCOM.SelectedIndex = 0
        Next

    End Sub

    Private Delegate Sub D_UpdateLabelText(ByVal label As Label, ByVal text As String)

    Private Sub setLabelText(label As Label, text As String)
        System.Threading.Thread.Sleep(1)
        label.Invoke(
            New D_UpdateLabelText(AddressOf UpdateLabelText), label, text)
    End Sub

    Private Sub UpdateLabelText(ByVal label As Label, ByVal text As String)
        label.Text = text
    End Sub

    Private Delegate Sub D_UpdateTextBoxText(ByVal Textbox As TextBox, ByVal text As String)

    Private Sub setTextBoxText(Textbox As TextBox, text As String)
        System.Threading.Thread.Sleep(1)
        Textbox.Invoke(
            New D_UpdateTextBoxText(AddressOf UpdateTextBoxText), Textbox, text)
    End Sub

    Private Sub UpdateTextBoxText(ByVal Textbox As TextBox, ByVal text As String)
        Textbox.Text = ""
        Textbox.SelectionStart = Textbox.Text.Length
        Textbox.SelectedText = text
    End Sub

    Private Delegate Sub D_UpdateButtonText(ByVal Button As Button, ByVal text As String)

    Private Sub setButtonText(Button As Button, text As String)
        System.Threading.Thread.Sleep(1)
        Button.Invoke(
            New D_UpdateButtonText(AddressOf UpdateButtonText), Button, text)
    End Sub

    Private Sub UpdateButtonText(ByVal Button As Button, ByVal text As String)
        Button.Text = text
    End Sub

    Private Sub WaitTm(ByVal intLen As Integer)
        Timer1.Interval = intLen
        Timer1.Start()
        Do While Timer1.Enabled = True
            Application.DoEvents()
        Loop

    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop()
    End Sub



    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        Timer2.Stop()
    End Sub



    Private Shared Sub Set_Data(ByRef buff As Byte, ByVal indata As Byte, ByRef pos As Byte)
        buff = indata
        pos += 1
    End Sub

    Private Shared Sub Set_Data(ByRef buff() As Byte, ByVal indata() As Byte, ByRef pos As Byte, ct As Byte)
        Dim i As Integer
        For i = 0 To ct - 1
            buff(i) = indata(pos)
            pos += 1
        Next
    End Sub

    Private Shared Sub Get_Data(ByRef buff As Byte, ByRef indata As Byte, ByRef len1 As Byte)
        indata = buff
        len1 += 1
    End Sub

    Private Shared Sub Get_Data(ByRef buff() As Byte, ByRef indata() As Byte, ByRef len1 As Byte, ct As Byte)
        For i = 0 To ct - 1
            indata(i) = buff(len1)
            len1 += 1
        Next
    End Sub

    Private Sub cmbTypeDo_SelectIndexChange(sender As Object, e As EventArgs)
        Dim cmb As ComboBox = TryCast(sender, ComboBox)
        If cmb IsNot Nothing Then
            Dim ItemValue As String = cmb.Text
            Debug.WriteLine(ItemValue)
            If (ItemValue.Contains("D2.Change FootPrint") = True) Then
                txtSourceRef.Enabled = True
                txtSpace.Enabled = False
                txtSpaceHight.Enabled = False
                txtMouseLocation(0).Enabled = False
                txtMouseLocation(1).Enabled = False
            ElseIf (ItemValue.Contains("D2.Change Name") = True) Then
                txtSourceRef.Enabled = True
                txtSpace.Enabled = False
                txtSpaceHight.Enabled = False
                txtMouseLocation(0).Enabled = False
                txtMouseLocation(1).Enabled = False
            ElseIf (ItemValue.Contains("D2.Hidden Value") = True) Then
                txtSourceRef.Enabled = False
                txtSpace.Enabled = False
                txtSpaceHight.Enabled = False
                txtMouseLocation(0).Enabled = False
                txtMouseLocation(1).Enabled = False
            ElseIf ItemValue.Contains("Al.MoveComp") = True Then
                txtSourceRef.Enabled = False
                txtSpace.Enabled = True
                txtSpaceHight.Enabled = True
                txtMouseLocation(0).Enabled = True
                txtMouseLocation(1).Enabled = True
            Else
                txtSourceRef.Enabled = False
                txtSpace.Enabled = False
                txtSpaceHight.Enabled = False
                txtMouseLocation(0).Enabled = False
                txtMouseLocation(1).Enabled = False
            End If
        End If
    End Sub

    'Processing button scan port 
    Private Sub btn_click(sender As Object, e As EventArgs)
        Dim btnClick As Button = TryCast(sender, Button)
        If btnClick IsNot Nothing Then
            'MsgBox(StrComName)
            Dim strDmy As String = btnClick.Name
            If strDmy.StartsWith("btnScanComPort") = True Then
                ScanComPort()
            ElseIf strDmy.StartsWith("btnHide") = True Then
                Debug.WriteLine("btnHide")
                NotifyIcon1.Visible = True
                Me.Hide()
                Me.ShowInTaskbar = True
                ' WindowState = FormWindowState.Minimized
                Debug.WriteLine("Hide")
            ElseIf strDmy.StartsWith("btnEnaDisCOM") = True Then
                Debug.WriteLine("btnEnaDisCOM")
                If InStr(1, btnEnaDisCOM.Text, "Opening", 1) Then
                    setButtonText(btnEnaDisCOM, "Closed COM")
                    serialPort.Close()
                Else
                    setButtonText(btnEnaDisCOM, "Opening COM")
                    serialPort.Open()
                End If
            ElseIf strDmy.StartsWith("btnScanApp") = True Then
                Debug.WriteLine("btnScanApp")
                ScanApp()
            ElseIf strDmy.StartsWith("btnAction") = True Then
                Activefunction(CInt(strDmy.Substring(strDmy.Length - 2, 2)))
            ElseIf strDmy.StartsWith("btnDo") = True Then
                Debug.WriteLine(strDmy)
                btnDo_Click()
            ElseIf strDmy.StartsWith("btnMouseLocation") = True Then
                btnMouseLocation_Click()
            ElseIf strDmy.StartsWith("btnFindNext") = True Then
                FindNext()
            ElseIf strDmy.StartsWith("btnFindBack") = True Then
                FindPre()
            Else
                Debug.WriteLine(strDmy)
            End If
        End If
    End Sub


    Private Sub Activefunction(No As Integer)
        setLabelText(lblStatus, "Send No: " & No)
        Dim strKey As String = ""
        If chkCtrl(No).Checked = True Then
            strKey &= Key.CTRL
        End If

        If chkShift(No).Checked = True Then
            strKey &= Key.SHIFT
        End If
        strKey &= cmbKey(No).Text

        Debug.WriteLine(cmbTarget.Text & ": " & strKey)
        If cmbTarget.SelectedIndex = 0 Then
            SendKey2App("", strKey, 1)
        Else
            SendKey2App(cmbTarget.Text, strKey, 1)
        End If
        WaitTm(500)
        setLabelText(lblStatus, "Free")
    End Sub

    Private Sub btnMouseLocation_Click()
        MsgBox("Start Get location Mouse")
        For i = 0 To 100
            setTextBoxText(txtMouseLocation(0), GetX().ToString)
            setTextBoxText(txtMouseLocation(1), GetY().ToString)
            Debug.WriteLine(GetX() & " - " & GetY())
            WaitTm(50)
        Next
        MsgBox("Finish")
    End Sub


    Private Sub btnDo_Click()
        setLabelText(lblStatus, "Changing")

        'cmbTypeDo.Text = "Al.MoveComp"
        'txtSource.Text = "SW807"


        'Check Ref
        Dim strKey As String = ""

        Key.init()
        If Control.IsKeyLocked(Keys.CapsLock) = True Then
            MsgBox("CapsLock is On")
            'GoTo ERROR_END
        End If

        'cmbTypeDo.Items.Add("D2.Change FootPrint")
        'cmbTypeDo.Items.Add("D2.Change Name")
        'cmbTypeDo.Items.Add("D2.Hidden Value")

        If cmbTypeDo.Text.Contains("D2.Change FootPrint") = True Then
            Dim result As DialogResult = MessageBox.Show("You want?", "D2.Change FootPrint", MessageBoxButtons.YesNo)
            If (result = DialogResult.OK) Then
            Else
                GoTo ERROR_END
            End If

            'Check txt source
            If txtSource.Text = "" Then
                GoTo ERROR_END
            End If
            'Check cmbDo
            If cmbTypeDo.SelectedIndex = 0 Then
                If txtSourceRef.Text = "" Then
                    GoTo ERROR_END
                End If
            ElseIf cmbTypeDo.SelectedIndex = 1 Then
                If txtSourceRef.Text = "" Then
                    GoTo ERROR_END
                End If
            End If

            'Debug.WriteLine(txtSource.Text)
            For Each data As String In txtSource.Text.Split(vbCrLf)
                data = data.Replace(vbCrLf, "")
                data = data.Replace(vbLf, "")
                data = data.Replace(vbLf, "")
                If cmbTarget.SelectedIndex = 0 Then
                    GoTo ERROR_END
                Else
                    If data <> "" Then
                        Debug.WriteLine(cmbTarget.Text & ": " & data)
                        'Active App D2 Cad
                        NextApp(cmbTarget.Text, 1)
                        'Find
                        strKey = Key.CTRL & Key.ff
                        If SendKey2App(cmbTarget.Text, strKey, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        'SetMousePos(850, 420)
                        'LeftClick()
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, data, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If

                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If

                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.ENTER, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        WaitTm(20)
                        LeftClick()
                        WaitTm(20)
                        LeftClick()
                        'Change Footprint
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.ENTER, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If


                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, "PCB", 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, txtSourceRef.Text, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.ENTER, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.ENTER, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.TAB, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                        If SendKey2App(cmbTarget.Text, Key.ENTER, 0) = False Then
                            MsgBox(data)
                            GoTo ERROR_END
                        End If
                    End If

                End If
                WaitTm(10)
            Next
        ElseIf cmbTypeDo.Text.Contains("D2.Change Name") = True Then
            MsgBox("Null")
        ElseIf cmbTypeDo.Text.Contains("D2.Hidden Value") = True Then
            MsgBox("Null")
        ElseIf cmbTypeDo.Text.Contains("Al.MoveComp") = True Then
            'Check txt source
            If txtSource.Text = "" Then
                GoTo ERROR_END
            End If
            'Check cmbDo
            If cmbTypeDo.SelectedIndex = 0 Then
                If txtSourceRef.Text = "" Then
                    GoTo ERROR_END
                End If
            ElseIf cmbTypeDo.SelectedIndex = 1 Then
                If txtSourceRef.Text = "" Then
                    GoTo ERROR_END
                End If
            End If

            'cmbTypeDo.Text = "Al.MoveComp"
            'txtSource.Text = "SW807"
            Dim poslen As Integer = 0
            Dim poshight As Integer = 0
            Key.init()

            For Each data As String In txtSource.Text.Split(vbCrLf)
                data = data.Replace(vbCrLf, "")
                data = data.Replace(vbLf, "")
                data = data.Replace(vbLf, "")
                If cmbTarget.SelectedIndex = 0 Then
                    GoTo ERROR_END
                Else
                    If data <> "" Then
                        'Active App D2 Cad
                        WaitTm(100)
                        NextApp(cmbTarget.Text, 0)
                        SetMousePos(931, 556)
                        WaitTm(10)
                        LeftClick()
                        WaitTm(50)
                        strKey = Key.CTRL & Key.ff
                        Try
                            SendKeys.Send(strKey)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)
                        Try
                            SendKeys.Send(data)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)
                        Try
                            SendKeys.Send(Key.ENTER)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(500)
                        LeftDown()
                        WaitTm(800)
                        SetMousePos(CInt(txtMouseLocation(0).Text) + poslen * CInt(txtSpace.Text), CInt(txtMouseLocation(1).Text) + CInt(txtSpaceHight.Text) * poshight)
                        poslen = poslen + 1
                        WaitTm(500)
                        LeftUp()
                        WaitTm(500)
                    Else
                        poslen = 0
                        poshight = poshight + 1
                    End If
                End If
            Next
        ElseIf cmbTypeDo.Text.Contains("RF007.TX-Idle") = True Then
            Key.init()
            WaitTm(100)
            While (True)
                NextApp(cmbTarget.Text, 0)
                strKey = Key.ENTER
                Try
                    SendKeys.Send(strKey)
                Catch ex As Exception
                    MsgBox("ERROR")
                    GoTo ERROR_END
                End Try
                WaitTm(GetRandom(1500, 3000))
            End While

        ElseIf cmbTypeDo.Text.Contains(FIND_FOOTPRINT) = True Then
            'Check txt source
            If txtSource.Text = "" Then
                GoTo ERROR_END
            End If
            Dim poslen As Integer = 0
            Dim poshight As Integer = 0
            Key.init()
            intFindFootprint = 0


#If 0 Then
            For Each data As String In txtSource.Text.Split(vbCrLf)
                data = data.Replace(vbCrLf, "")
                data = data.Replace(vbLf, "")
                data = data.Replace(vbLf, "")
                If cmbTarget.SelectedIndex = 0 Then
                    GoTo ERROR_END
                Else
                    If data <> "" Then
                        WaitTm(100)
                        NextApp(cmbTarget.Text, 0)
                        SetMousePos(931, 556)
                        WaitTm(10)
                        LeftClick()
                        WaitTm(50)
                        strKey = Key.CTRL & Key.ff
                        Try
                            SendKeys.Send(strKey)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)
                        Try
                            SendKeys.Send(data)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)

                        Try
                            SendKeys.Send(Key.ENTER)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try

                        WaitTm(500)
                    Else
                        poslen = 0
                        poshight = poshight + 1
                    End If
                End If
            Next
#End If
        End If

        WaitTm(500)
        'MsgBox("Finish")
        setLabelText(lblStatus, "Free")
        Exit Sub
ERROR_END:

        WaitTm(500)
        setLabelText(lblStatus, "Error")
        Exit Sub
    End Sub

    Private Sub FindNext()
        'Check txt source
        If txtSource.Text = "" Then
            GoTo ERROR_END
        End If
        Dim poslen As Integer = 0
        Dim poshight As Integer = 0
        Dim strData As String = ""
        Key.init()
        Dim strKey As String = ""
        For Each data As String In txtSource.Text.Split(vbCrLf)
            data = data.Replace(vbCrLf, "")
            data = data.Replace(vbLf, "")
            data = data.Replace(vbLf, "")
            If cmbTarget.SelectedIndex = 0 Then
                GoTo ERROR_END
            Else
                If (poslen < intFindFootprint) Then
                    poslen = poslen + 1
                Else
                    If data <> "" Then
                        WaitTm(100)
                        NextApp(cmbTarget.Text, 0)
                        SetMousePos(931, 556)
                        WaitTm(10)
                        LeftClick()
                        WaitTm(50)
                        strKey = Key.CTRL & Key.ff
                        Try
                            SendKeys.Send(strKey)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)
                        Try
                            SendKeys.Send(data)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)

                        Try
                            SendKeys.Send(Key.ENTER)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        strData = data
                        WaitTm(500)
                    End If

                    GoTo WORK_END
                End If
            End If
        Next

WORK_END:
        intFindFootprint = intFindFootprint + 1
        'MsgBox(strData)
ERROR_END:

    End Sub

    Private Sub FindPre()
        'Check txt source
        If txtSource.Text = "" Then
            GoTo ERROR_END
        End If
        Dim poslen As Integer = 0
        Dim poshight As Integer = 0
        Key.init()
        Dim strKey As String = ""
        Dim strData As String = ""

        For Each data As String In txtSource.Text.Split(vbCrLf)
            data = data.Replace(vbCrLf, "")
            data = data.Replace(vbLf, "")
            data = data.Replace(vbLf, "")
            If cmbTarget.SelectedIndex = 0 Then
                GoTo ERROR_END
            Else
                If (poslen < intFindFootprint) Then
                    poslen = poslen + 1
                Else
                    If data <> "" Then
                        WaitTm(100)
                        NextApp(cmbTarget.Text, 0)
                        SetMousePos(931, 556)
                        WaitTm(10)
                        LeftClick()
                        WaitTm(50)
                        strKey = Key.CTRL & Key.ff
                        Try
                            SendKeys.Send(strKey)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)
                        Try
                            SendKeys.Send(data)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        WaitTm(100)

                        Try
                            SendKeys.Send(Key.ENTER)
                        Catch ex As Exception
                            MsgBox(data)
                            GoTo ERROR_END
                        End Try
                        strData = data
                        WaitTm(500)
                    End If

                    GoTo WORK_END
                End If
            End If
        Next

WORK_END:
        If (intFindFootprint > 0) Then
            intFindFootprint = intFindFootprint - 1
        End If
        'MsgBox(strData)


ERROR_END:

    End Sub

    Public Function GetRandom(ByVal Min As Integer, ByVal Max As Integer) As Integer
        Dim Generator As System.Random = New System.Random()
        Return Generator.Next(Min, Max)
    End Function

    Function ShiftKeyData(data As String, blUp As Boolean) As String
        If Control.IsKeyLocked(Keys.CapsLock) Then
            'MessageBox.Show("The Caps Lock key is ON.")
            If (blUp = True) Then
                ShiftKeyData = data
            Else
                ShiftKeyData = "+" & data
            End If
        Else
            'MessageBox.Show("The Caps Lock key is OFF.")
            If (blUp = True) Then
                ShiftKeyData = "+" & data
            Else
                ShiftKeyData = data
            End If
        End If
    End Function


    Private Function SendKey2App(apps As String, key As String, ModeRe As Integer) As Boolean
        SendKey2App = True
        'ModeRe = 0 'A # a
        ModeRe = 1 'A = a
        'ModeRe = 2 'Other
        WaitTm(2)
        If ModeRe < 3 Then
            'Scan App
            Dim ImgList As New ImageList
            Dim ProcID As Integer = 0
            For Each proc As Process In Process.GetProcesses
                If proc.MainWindowTitle <> "" Then
                    'ImgList.Images.Add(Icon.ExtractAssociatedIcon(proc.MainModule.FileName))
                    'Dim lvi As New ListViewItem(proc.ProcessName, ImgList.Images.Count - 1)
                    'lvi.SubItems.Add(proc.MainModule.FileName)
                    'Debug.WriteLine(lvi)

                    If InStr(1, proc.ProcessName, apps, ModeRe) Then
                        'Debug.WriteLine(proc.Id)
                        'Debug.WriteLine("ok")
                        'Debug.WriteLine(proc.ProcessName)
                        'SendKeys.Send("^+{TAB}")
                        ProcID = proc.Id
                    End If
                End If
            Next

            'Send key
            If apps = "" Then
                SendKeys.Send(key)
            Else
                If ProcID <> 0 Then
                    Try
                        'AppActivate(ProcID)
                        SendKeys.Send(key)
                    Catch ex As Exception
                        SendKey2App = False
                    End Try
                End If
            End If

        Else
            Debug.WriteLine("Mode Referent NG")
        End If
    End Function

    Private Sub NextApp(apps As String, ModeRe As Integer)
        'ModeRe = 0 'A # a
        ModeRe = 1 'A = a
        'ModeRe = 2 'Other

        If ModeRe < 3 Then
            'Scan App
            Dim ImgList As New ImageList
            Dim ProcID As Integer = 0
            For Each proc As Process In Process.GetProcesses
                If proc.MainWindowTitle <> "" Then
                    'ImgList.Images.Add(Icon.ExtractAssociatedIcon(proc.MainModule.FileName))
                    'Dim lvi As New ListViewItem(proc.ProcessName, ImgList.Images.Count - 1)
                    'lvi.SubItems.Add(proc.MainModule.FileName)
                    'Debug.WriteLine(lvi)

                    If InStr(1, proc.ProcessName, apps, ModeRe) Then
                        'Debug.WriteLine(proc.Id)
                        'Debug.WriteLine("ok")
                        'Debug.WriteLine(proc.ProcessName)
                        'SendKeys.Send("^+{TAB}")
                        ProcID = proc.Id
                    End If
                End If
            Next

            'Send key
            If ProcID <> 0 Then
                AppActivate(ProcID)
            End If

        Else
            Debug.WriteLine("Mode Referent NG")
        End If
    End Sub

    Private Sub ScanApp()
        Dim ImgList As New ImageList
        Dim intDmy As Integer = 0
        intDmy = cmbTarget.SelectedIndex
        txtList.Text = ""
        cmbTarget.Items.Clear()
        cmbTarget.Items.Add("App Activing")

        For Each proc As Process In Process.GetProcesses
            If proc.MainWindowTitle <> "" Then
                'ImgList.Images.Add(Icon.ExtractAssociatedIcon(proc.MainModule.FileName))
                ' Dim lvi As New ListViewItem(proc.ProcessName, ImgList.Images.Count - 1)
                'lvi.SubItems.Add(proc.MainModule.FileName)
                'txtList.Text &= lvi.Text & vbNewLine
                'cmbTarget.Items.Add(lvi.Text)
                cmbTarget.Items.Add(proc.ProcessName)
                txtList.Text &= proc.ProcessName & vbNewLine
            End If
        Next
        cmbTarget.SelectedIndex = intDmy
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.ShowInTaskbar = True
        Me.Show()
        NotifyIcon1.Visible = False
        Debug.WriteLine("Show")
    End Sub

    '
    'At load
    '
    Private Sub Form_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ScanComPort()

        NotifyIcon1.Visible = False
        Debug.WriteLine("MyBase.Load")
        Dim strPath As String = Application.StartupPath()
        If System.IO.File.Exists(strPath & "\" & "Boot.TriggerKey") Then
            ReadFile(strPath, "Boot.TriggerKey")
        End If
        If (cmbCOM.Items.Count > 0) Then
            serialPort = New SerialPort With {.PortName = cmbCOM.Text, .BaudRate = 115200, .Parity = 0, .DataBits = 8, .StopBits = 1}
        End If


        'LeftClick()
        'RightClick()
        'SetMousePos(100, 100)
    End Sub

    Private Sub ReadFile(strPath As String, namefile As String)
        Dim Read As StreamReader
        Dim strLine As String = ""
        'Dim strDmy As String = ""

        Read = New StreamReader(strPath & "\" & namefile)

        'Read name
        strLine = Read.ReadLine

        'Read ComPort
        Dim blComCheck As Boolean = False
        strLine = Read.ReadLine
        'Debug.WriteLine(strLine)
        For i = 0 To cmbCOM.Items.Count - 1
            cmbCOM.SelectedIndex = i
            If InStr(1, strLine, cmbCOM.Text, 1) Then
                cmbCOM.SelectedIndex = i
                blComCheck = True
                Exit For
            End If
        Next
        If blComCheck = False Then
            If (cmbCOM.Items.Count > 0) Then
                cmbCOM.SelectedIndex = 0
            End If
        End If

        'Read key set
        strLine = Read.ReadLine
        Do While Not strLine Is Nothing
            If CheckStrIsNum(strLine.Substring(5, 2)) Then
                Dim no As Integer = CInt(strLine.Substring(5, 2))
                'Debug.WriteLine(no)
                Dim strDmy As String
                'Ctrl
                If CheckStrIsNum(strLine.Substring(8, 1)) Then
                    strDmy = strLine.Substring(8, 1)
                    chkCtrl(no).Checked = IIf(strDmy = "1", True, False)
                End If
                'Shift
                If CheckStrIsNum(strLine.Substring(9, 1)) Then
                    strDmy = CInt(strLine.Substring(9, 1))
                    chkShift(no).Checked = IIf(strDmy = "1", True, False)
                End If
                'Key 
                If CheckStrIsNum(strLine.Substring(10, 2)) Then
                    cmbKey(no).SelectedIndex = CInt(strLine.Substring(10, 2))
                End If
            End If
            strLine = Read.ReadLine
        Loop

        'Close the file.
        Read.Close()
    End Sub

    Private Sub WriteFile(strPath As String, namefile As String)
        Dim Writer As StreamWriter

        Dim strLine As String = ""
        Dim strDmy As String = ""

        'Debug.WriteLine(strPath)
        Writer = New StreamWriter(strPath & "\" & namefile)

        'Write line
        Writer.WriteLine("SendKey_Ver " & VER)

        Writer.WriteLine("COM:" & cmbCOM.Text)
        For i = 0 To chkCtrl.Count - 1
            Writer.WriteLine("KEYNO" & Format(i, "00") & ":" & IIf(chkCtrl(i).Checked, 1, 0) & IIf(chkShift(i).Checked, 1, 0) & Format(cmbKey(i).SelectedIndex, "00"))
        Next

        Writer.Close()
    End Sub



    Function CheckStrIsNum(str As String) As Boolean
        CheckStrIsNum = True
        For i = 0 To str.Length - 1
            If InStr(1, "0123456789", str.Substring(i, 1), 1) = False Then
                CheckStrIsNum = False
                Exit For
            End If
        Next
    End Function

    '
    'At Close
    '
    Private Sub Form_Close(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        Dim strPath As String = Application.StartupPath()
        WriteFile(strPath, "Boot.TriggerKey")
        NotifyIcon1.Visible = False
        Debug.WriteLine("MyBase.Closed")
        SerialPort.Close()

    End Sub


    Private Structure stKey
        Dim BACKSPACE As String
        Dim BREAK As String
        Dim APSLOCK As String
        Dim DELETE As String
        Dim DOWN As String
        Dim KEND As String
        Dim ENTER As String
        Dim ESC As String
        Dim HELP As String
        Dim HOME As String
        Dim INSERT As String
        Dim LEFT As String
        Dim NUMLOCK As String
        Dim PAGEDOWN As String
        Dim PAGEUP As String
        Dim PRINTSCREEN As String
        Dim RIGHT As String
        Dim SCROLLLOCK As String
        Dim TAB As String
        Dim UP As String
        Dim F1 As String
        Dim F2 As String
        Dim F3 As String
        Dim F4 As String
        Dim F5 As String
        Dim F6 As String
        Dim F7 As String
        Dim F8 As String
        Dim F9 As String
        Dim F10 As String
        Dim F11 As String
        Dim F12 As String
        Dim F13 As String
        Dim F14 As String
        Dim F15 As String
        Dim F16 As String
        Dim F17 As String
        Dim F18 As String
        Dim F19 As String
        Dim F20 As String
        Dim F21 As String
        Dim F22 As String
        Dim F23 As String
        Dim F24 As String

        Dim ADD As String
        Dim SUBTRACT As String
        Dim MULTIPLY As String
        Dim DIVIDE As String

        Dim A As String
        Dim B As String
        Dim C As String
        Dim D As String
        Dim E As String
        Dim F As String
        Dim G As String
        Dim H As String
        Dim I As String
        Dim J As String
        Dim K As String
        Dim L As String
        Dim M As String
        Dim N As String
        Dim O As String
        Dim P As String
        Dim Q As String
        Dim R As String
        Dim S As String
        Dim T As String
        Dim U As String
        Dim V As String
        Dim W As String
        Dim X As String
        Dim Y As String
        Dim Z As String

        Dim aa As String
        Dim bb As String
        Dim cc As String
        Dim dd As String
        Dim ee As String
        Dim ff As String
        Dim gg As String
        Dim hh As String
        Dim ii As String
        Dim jj As String
        Dim kk As String
        Dim ll As String
        Dim mm As String
        Dim nn As String
        Dim oo As String
        Dim pp As String
        Dim qq As String
        Dim rr As String
        Dim ss As String
        Dim tt As String
        Dim uu As String
        Dim vv As String
        Dim ww As String
        Dim xx As String
        Dim yy As String
        Dim zz As String

        Dim N0 As String
        Dim N1 As String
        Dim N2 As String
        Dim N3 As String
        Dim N4 As String
        Dim N5 As String
        Dim N6 As String
        Dim N7 As String
        Dim N8 As String
        Dim N9 As String

        Dim SHIFT As String
        Dim CTRL As String
        Dim ALT As String

        Sub init()
            BACKSPACE = "{BACKSPACE}" ', {BS}, or {BKSP}
            BREAK = "{BREAK}"
            APSLOCK = "{CAPSLOCK}"
            DELETE = "{DELETE}" ' or {DEL}
            DOWN = "{DOWN}"
            KEND = "{END}"
            ENTER = "{ENTER}" '{ENTER}or ~
            ESC = "{ESC}"
            HELP = "{HELP}"
            HOME = "{HOME}"
            INSERT = "{INSERT}" '{INSERT} or {INS}
            LEFT = "{LEFT}"
            NUMLOCK = "{NUMLOCK}"
            PAGEDOWN = "{PGDN}"
            PAGEUP = "{PGUP}"
            PRINTSCREEN = "{PRTSC}"
            RIGHT = "{RIGHT}"
            SCROLLLOCK = "{SCROLLLOCK}"
            TAB = "{TAB}"
            UP = "{UP}"
            F1 = "{F1}"
            F2 = "{F2}"
            F3 = "{F3}"
            F4 = "{F4}"
            F5 = "{F5}"
            F6 = "{F6}"
            F7 = "{F7}"
            F8 = "{F8}"
            F9 = "{F9}"
            F10 = "{F10}"
            F11 = "{F11}"
            F12 = "{F12}"
            F13 = "{F13}"
            F14 = "{F14}"
            F15 = "{F15}"
            F16 = "{F16}"
            F17 = "{F17}"
            F18 = "{F18}"
            F19 = "{F19}"
            F20 = "{F20}"
            F21 = "{F21}"
            F22 = "{F22}"
            F23 = "{F23}"
            F24 = "{F24}"
            ADD = "{ADD}"
            SUBTRACT = "{SUBTRACT}"
            MULTIPLY = "{MULTIPLY}"
            DIVIDE = "{DIVIDE}"

            A = "{A}"
            B = "{B}"
            C = "{C}"
            D = "{D}"
            E = "{E}"
            F = "{F}"
            G = "{G}"
            H = "{H}"
            I = "{I}"
            J = "{J}"
            K = "{K}"
            L = "{L}"
            M = "{M}"
            N = "{N}"
            O = "{O}"
            P = "{P}"
            Q = "{Q}"
            R = "{R}"
            S = "{S}"
            T = "{T}"
            U = "{U}"
            V = "{V}"
            W = "{W}"
            X = "{X}"
            Y = "{Y}"
            Z = "{Z}"

            aa = "{a}"
            bb = "{b}"
            cc = "{c}"
            dd = "{d}"
            ee = "{e}"
            ff = "{f}"
            gg = "{g}"
            hh = "{h}"
            ii = "{i}"
            jj = "{j}"
            kk = "{k}"
            ll = "{l}"
            mm = "{m}"
            nn = "{n}"
            oo = "{o}"
            pp = "{p}"
            qq = "{q}"
            rr = "{r}"
            ss = "{s}"
            tt = "{t}"
            uu = "{u}"
            vv = "{v}"
            ww = "{w}"
            xx = "{x}"
            yy = "{y}"
            zz = "{z}"

            N0 = "{0}"
            N1 = "{1}"
            N2 = "{2}"
            N3 = "{3}"
            N4 = "{4}"
            N5 = "{5}"
            N6 = "{6}"
            N7 = "{7}"
            N8 = "{8}"
            N9 = "{9}"

            CTRL = "^"
            SHIFT = "+"
            ALT = "%"
        End Sub
    End Structure
    Private Key As stKey

    Public Delegate Sub StringSubPointer(ByVal Buffer As String)

    Private Sub serialPortGetData(ByVal Buffer As String)
        'Received.AppendText(Buffer)
        Debug.WriteLine(Buffer)
        '------------------------------------------------------------------------
        'Receiver Serial at here
        If InStr(1, Buffer, "FI~TP0000", 0) Then
            Activefunction(CInt(Buffer.Substring(Buffer.Length - 2, 2)))
        End If
        'Activefunction(0)
    End Sub

    Private Sub Receiver(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs) Handles serialPort.DataReceived
        Me.BeginInvoke(New StringSubPointer(AddressOf serialPortGetData), serialPort.ReadLine)
    End Sub

End Class



Module Module1
    Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
    Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
    Public Declare Function GetCursorPos Lib "user32" (ByVal lpPoint As POINTAPI) As Long

    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4
    Public Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Public Const MOUSEEVENTF_MIDDLEUP = &H40
    Public Const MOUSEEVENTF_RIGHTDOWN = &H8
    Public Const MOUSEEVENTF_RIGHTUP = &H10
    Public Const MOUSEEVENTF_MOVE = &H1

    Public Structure POINTAPI
        Dim x As Integer
        Dim y As Integer
    End Structure

    Public Function GetX() As Integer
        'Dim n As POINTAPI
        'GetCursorPos(n)
        'GetX = n.x
        'Debug.WriteLine(n.x)
        Dim MousePosition As Point
        MousePosition = Cursor.Position
        'Debug.WriteLine(MousePosition.X)
        GetX = MousePosition.X
    End Function

    Public Function GetY() As Integer
        'Dim n As POINTAPI
        'GetCursorPos(n)
        Dim MousePosition As Point
        MousePosition = Cursor.Position
        'Debug.WriteLine(MousePosition.Y)
        GetY = MousePosition.Y
    End Function

    Public Sub LeftClick()
        LeftDown()
        LeftUp()
    End Sub

    Public Sub LeftDown()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End Sub

    Public Sub LeftUp()
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

    Public Sub MiddleClick()
        MiddleDown()
        MiddleUp()
    End Sub

    Public Sub MiddleDown()
        mouse_event(MOUSEEVENTF_MIDDLEDOWN, 0, 0, 0, 0)
    End Sub

    Public Sub MiddleUp()
        mouse_event(MOUSEEVENTF_MIDDLEUP, 0, 0, 0, 0)
    End Sub

    Public Sub RightClick()
        RightDown()
        RightUp()
    End Sub

    Public Sub RightDown()
        mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0)
    End Sub

    Public Sub RightUp()
        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0)
    End Sub

    Public Sub MoveMouse(ByVal xMove As Long, ByVal yMove As Long)
        mouse_event(MOUSEEVENTF_MOVE, xMove, yMove, 0, 0)
    End Sub

    Public Sub SetMousePos(ByVal xPos As Long, ByVal yPos As Long)
        SetCursorPos(xPos, yPos)
    End Sub
End Module
