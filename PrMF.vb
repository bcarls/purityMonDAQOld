Option Strict Off
Option Explicit On
Imports Microsoft.VisualBasic.PowerPacks
Friend Class PrMF '
    Inherits System.Windows.Forms.Form

    ' Win32 exports
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)
    Private Declare Function GetTickCount Lib "kernel32" () As Integer
    Private Declare Function GetInputState Lib "user32" () As Integer

    Private Declare Function BmpToJpeg Lib "Bmp2Jpeg.dll" (ByVal BmpFilename As String, ByVal JpegFilename As String, ByVal CompressQuality As Short) As Short

    ' Input range identifiers
    Dim InputRangeIdChA As Integer
    Dim InputRangeIdChB As Integer

    'Create the Graphics object
    Dim g As Graphics = Me.CreateGraphics

    Private Const PI As Single = 3.14159265
    Private Const DALPHA As Single = PI / 6
    Private Const Data7On As Short = 128
    Private Const HVOn As Short = 16
    Private m_Alpha As Single
    Private Sub CirclePoint(ByVal theta As Single, ByVal circle_radius As Single, ByRef X As Single, ByRef Y As Single)
        X = circle_radius * System.Math.Cos(theta)
        Y = circle_radius * System.Math.Sin(theta)
    End Sub
    Private Sub DrawOrbitEllipse(ByVal radius As Single, ByVal sx As Single, ByVal sy As Single, ByVal angle As Single, ByVal tx As Single, ByVal ty As Single, ByVal clr As System.Drawing.Color)
        Const DTHETA As Single = PI / 20
        Dim x0 As Single
        Dim y0 As Single
        Dim X As Single
        Dim Y As Single
        Dim CurrentX As Single
        Dim CurrentY As Single
        Dim theta As Single
        Dim ellipsePen As New System.Drawing.Pen(clr)

        ' Get the first point.
        EllipsePoint(X, Y, 0, radius, sx, sy, angle, tx, ty)
        'UPGRADE_ISSUE: Form property PrMF.CurrentX was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        CurrentX = X
        'UPGRADE_ISSUE: Form property PrMF.CurrentY was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        CurrentY = Y
        x0 = X
        y0 = Y

        ' Draw other points.
        For theta = DTHETA To 2 * PI Step DTHETA
            CurrentX = X
            CurrentY = Y
            EllipsePoint(X, Y, theta, radius, sx, sy, angle, tx, ty)
            'UPGRADE_ISSUE: Form method PrMF.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
            'Me.Line (X, Y), clr
            g.DrawLine(ellipsePen, CurrentX, CurrentY, X, Y)
        Next theta

        ' Close the circle.
        'UPGRADE_ISSUE: Form method PrMF.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Line (x0, y0), clr
        g.DrawLine(ellipsePen, CurrentX, CurrentX, x0, y0)
    End Sub
    Private Sub EllipsePoint(ByRef X As Single, ByRef Y As Single, ByVal theta As Single, ByVal radius As Single, ByVal sx As Single, ByVal sy As Single, ByVal angle As Single, ByVal tx As Single, ByVal ty As Single)
        CirclePoint(theta, radius, X, Y)
        ScalePoint(sx, sy, X, Y)
        RotatePoint(angle, X, Y)
        TranslatePoint(tx, ty, X, Y)
    End Sub
    Private Sub RotatePoint(ByVal theta As Single, ByRef X As Single, ByRef Y As Single)
        Dim new_x As Single
        Dim new_y As Single

        new_x = X * System.Math.Cos(theta) + Y * System.Math.Sin(theta)
        new_y = X * System.Math.Sin(theta) - Y * System.Math.Cos(theta)
        X = new_x
        Y = new_y
    End Sub
    Private Sub ScalePoint(ByVal scale_x As Single, ByVal scale_y As Single, ByRef X As Single, ByRef Y As Single)
        X = X * scale_x
        Y = Y * scale_y
    End Sub
    Private Sub TranslatePoint(ByVal tx As Single, ByVal ty As Single, ByRef X As Single, ByRef Y As Single)
        X = X + tx
        Y = Y + ty
    End Sub

    'UPGRADE_WARNING: Event Check2.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Check2_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check2.CheckStateChanged
        Dim Index As Short = Check2.GetIndex(eventSender)
        'Check2(Index).Value = 1
        Check3(Index).CheckState = System.Windows.Forms.CheckState.Unchecked
        If Index = 0 Then Ichan1 = 4
        If Index = 1 Then Ichan2 = 5
        If Index = 2 Then Ichan3 = 6

    End Sub

    'UPGRADE_WARNING: Event Check3.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Check3_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check3.CheckStateChanged
        Dim Index As Short = Check3.GetIndex(eventSender)
        Check2(Index).CheckState = System.Windows.Forms.CheckState.Unchecked
        'Check3(Index).Value = 1
        If Index = 0 Then Ichan1 = 0
        If Index = 1 Then Ichan2 = 1
        If Index = 2 Then Ichan3 = 2

    End Sub

    'UPGRADE_WARNING: Event Check5.CheckStateChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Check5_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Check5.CheckStateChanged
        Dim i As Object
        If Check5.CheckState = 1 Then
            '    If IOk = 0 Then IOk = 1
            For i = 0 To NPrM - 1
                Check4(i).Enabled = True
                Check4(i).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            Next i
        End If
    End Sub



    'UPGRADE_WARNING: Event Combo1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'UPGRADE_WARNING: ComboBox event Combo1.Change was upgraded to Combo1.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    Private Sub Combo1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo1.TextChanged
        ' InputRangeIdChA = CInt(VB6.GetItemString(Combo1, Combo1.SelectedIndex))
        InputRangeIdChA = VB6.GetItemData(Combo1, Combo1.SelectedIndex)

        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        Dim result As Boolean
        result = ConfigureBoard(boardHandle)
        If (result <> True) Then
            Exit Sub
        End If
    End Sub

    'UPGRADE_WARNING: Event Combo2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    'UPGRADE_WARNING: ComboBox event Combo2.Change was upgraded to Combo2.TextChanged which has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="DFCDE711-9694-47D7-9C50-45A99CD8E91E"'
    Private Sub Combo2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Combo2.TextChanged
        'InputRangeIdChB = CInt(VB6.GetItemString(Combo2, Combo2.SelectedIndex))
        InputRangeIdChB = VB6.GetItemData(Combo2, Combo2.SelectedIndex)

        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        Dim result As Boolean
        result = ConfigureBoard(boardHandle)
        If (result <> True) Then
            Exit Sub
        End If
    End Sub

    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        NullCmd.Focus()
        If iRunning = 1 Then Exit Sub
        If Command1.Text = "DAQ Stopped - Press to Start DAQ" Then
            Command1.Text = "DAQ Running - Press to Stop"
            StatusL.Text = "DAQ Started"
            xRate = 1 'cause to do one right now
            NullCmd.Focus()
            FileOpen(88, DataFileNameHist, OpenMode.Append)
            PrintLine(88, "+++++++" & Today & " " & TimeOfDay & " DAQ Started")
            FileClose(88)
        Else
            Command1.Text = "DAQ Stopped - Press to Start DAQ"
            StatusL.Text = "DAQ Stopped"
            NullCmd.Focus()
            FileOpen(88, DataFileNameHist, OpenMode.Append)
            PrintLine(88, "-------" & Today & " " & TimeOfDay & " DAQ Stopped")
            FileClose(88)
        End If
    End Sub


    Private Sub Command4_Click()

    End Sub

    'Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
    '	Out(Val("&H37a"), 15)
    '	FileOpen(3, xPath & "Run_num.ini", OpenMode.Input)
    '	Input(3, iiRun)
    '	iiRun = iiRun
    '	FileClose(3)
    '	FileOpen(3, xPath & "Run_num.ini", OpenMode.Output)
    '	PrintLine(3, iiRun)
    '	FileClose(3)
    '	RunNumL.Text = VB6.Format(iiRun, "######")
    '	DataFileName = DataFilePath & "Run_" & VB6.Format(iiRun, "000000") & ".txt"
    '	RunFileL.Text = DataFileName
    '	'ScopeWait.Enabled = True
    '	Out(Val("&H37a"), 0)

    'End Sub

    Private Sub Command3_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command3.Click
        Me.PrintForm1.Print(Me, PowerPacks.Printing.PrintForm.PrintOption.CompatibleModeClientAreaOnly)
    End Sub



    'Private Sub DAQDetails_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles DAQDetails.Click
    '		Hist.Visible = True
    '	End Sub

    Private Sub PrMF_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.DoubleClick
        Me.Height = VB6.TwipsToPixelsY(9630)
    End Sub

    Private Sub PrMF_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim xxx As Object
        Dim zppp As Object
        Dim xinc As Object
        'Dim i As Object
        Dim zztime As Object
        Me.Text = "PRM Data Acquisition Software Ver 4.1 AEB TY BC PRM v17" '   by A Baumbaugh, T Yang, and B Carls, Fermilab"
        ' 4.2  Now plotting the averaged and smoothed traces in the window
        ' 4.1  Configuring the program for uBooNE, has Emily's correction and removes the 1.05
        ' 4.0  Overhauling code to handle input from Alazar PCI digitizer model ATS310
        ' 3.05 set vscale=20mV for inline PrM, save image to ifix
        ' 3.04 take difference between ch2 and ch1 as cathode signal to reduce noise, save logfile to dropbox
        ' 3.03 change start time for anode from 25 us to 100 us
        ' 3.02 use different time scales for different monitors, save output files to dropbox
        ' 3.01 add support for liquid level interlock
        ' 3.00 add support for mux - T.Yang
        ' 2.92 chnged defaults for runs and irate
        ' 2.91 added a set focus to nullcmd to start and stop DAQ and added local data csv
        ' 2.9  changed default on channel 3 to RAW and default path to E:\ and added long history file
        ' 2.8  ADDED Luke.PRM_IMPURITIES.F_CV, TO OUTPUT FILE
        ' 2.7  changed csv file to match desired format
        ' 2.6  added .csv file writing to DAQ
        ' 2.5  made 30 sec start delay
        ' 2.4  fixed array too big problem
        ' 2.3  added 10 sec wait on text change and print button and fixed irunning error
        ' 2.2  fixed going negative problem or at least added debug to find it
        ' 2.1  added variable number of data sets
        ' 2.0  added changeable sense of liquid level and diode peak to condensed file
        ' 1.E  Further changes in analysis to avoid /0
        ' 1.D  further changes
        ' 1.C  added display of results to top form in a list box
        ' 1.B  fixed turning pulser on/off problem
        ' 1.A  fixed saving same name snafu
        ' 1.9  fixed problem caused by smaller settings form
        ' 1.8  shifted search for pulses to 25 us after trig
        ' 1.7  fancier argon, changed math again
        ' 1.6  fixed Argon and variable color fnal
        ' 1.5  modified algorithm
        ' 1.4  modified smoothing and fixed .txt files etc
        ' 1.3  added analysis code and printing
        ' 1.1  added smoothing and columnar print out
        ' 1.2  added scale divisors and changed smoothing
        'Private Sub Command1_Click()
        'Text2.Text = Str(Inp(Val("&H" + Text1.Text)))
        'End Sub
        '
        'Private Sub Command2_Click()
        'Out Val("&H" + Text1.Text), Val(Text2.Text)
        'End Sub
        'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        zztime = VB6.Format(TimeOfDay, "hh:mm:ss")
        '''''''''''''''''''''''''''''''
        'LogPath = "E:\"
        'LogPath = "C:\TJTest\"
        LogPath = "C:\PrM Log\"
        '''''''''''''''''''''''''''''''
        'ImagePath = "F:\"
        ImagePath = "C:\PrM Images\"
        'OneTrueLiquid = 1
        ISets = 1
        IOk = 0
        'Label1(1).ToolTipText = "Low is good...Dbl Click to Change"
        RMSCut = 10.0#
        'Taking out vScale, doesn't seem to be used at all
        'For i = 1 To 40
        'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'vScale(i) = 1
        'Next i
        ISmooth = 40
        IRate = 180
        xRate = IRate * 60
        'UPGRADE_WARNING: Couldn't resolve default property of object xinc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xinc = 1
        IPrM = 0
        NPrM = 5
        'UPGRADE_WARNING: Couldn't resolve default property of object zlocal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object zppp. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        zppp = CurDir() & "\"
        StatusL.Text = "DAQ Stopped"
        xPath = "C:\PrM1\"
        '''''''''''''''''''''''''''''''''''
        DataFilePath = "C:\PrM Data\"
        'DataFilePath = "C:\TJTest\"
        '''''''''''''''''''''''''''''''''''

        AllTracesPath = "C:\Users\bcarls\Desktop\All PrM traces\"

        DataFileNameHist = DataFilePath & "PrMLongHistory.txt"
        Label2(1).BringToFront()
        RunFileL.Text = xPath

        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        'UPGRADE_WARNING: Couldn't resolve default property of object zlocal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (boardHandle = 0) Then
            ' Oops, the device isn't available, notify the user and exit
            MsgBox("BoardID " & boardID & " not available." & Chr(10) & "You need to select an available device" & Chr(10) & "for the Alazar boardID in the VB IDE.", MsgBoxStyle.OkOnly, "Startup Error")
            End
        End If


        'Out Val("&H37a"), 0 'set output low
        Out(Val("&H378"), 8 + Data7On) ' turn off
        FileOpen(77, xPath & "settings.ini", OpenMode.Input)
        Do Until EOF(77)
            xxx = LineInput(77)
            'UPGRADE_WARNING: Couldn't resolve default property of object xxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xxx = UCase(xxx)
            'UPGRADE_WARNING: Couldn't resolve default property of object xxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(1, xxx, "ATOM ON") <> 0 Then IAtom = 1
            'UPGRADE_WARNING: Couldn't resolve default property of object xxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If InStr(1, xxx, "ATOM OFF") <> 0 Then IAtom = 0
            '    If InStr(1, xxx, "TIMEOUT 1 ON") <> 0 Then
            '        IgnorTimeOuts(1) = 0
            '        ix = InStr(1, xxx, "=")
            '        If ix <> 0 Then
            '            zzz = Mid(xxx, ix + 1)
            '            xdelay = Val(zzz)
            '            PMTTimeOut(1) = xdelay
            '            If PMTTimeOut(1) > 3000 Then PMTTimeOut(1) = 3000
            '        Else
            '            PMTTimeOut(1) = 100
            '        End If
            '    End If
        Loop
        FileClose(77)
        Ichan1 = 4 'default smoothed
        Ichan2 = 5 'default smoothed
        Ichan3 = 2 'default raw
        List1.Height = VB6.TwipsToPixelsY(1620)
        List2.Height = VB6.TwipsToPixelsY(1820)



        Combo1.Items.Add(New VB6.ListBoxItem("+/- 40 mV", INPUT_RANGE_PM_40_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 50 mV", INPUT_RANGE_PM_50_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 80 mV", INPUT_RANGE_PM_80_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 100 mV", INPUT_RANGE_PM_100_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 200 mV", INPUT_RANGE_PM_200_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 400 mV", INPUT_RANGE_PM_400_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 500 mV", INPUT_RANGE_PM_500_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 800 mV", INPUT_RANGE_PM_800_MV))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 1 V", INPUT_RANGE_PM_1_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 2 V", INPUT_RANGE_PM_2_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 4 V", INPUT_RANGE_PM_4_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 5 V", INPUT_RANGE_PM_5_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 8 V", INPUT_RANGE_PM_8_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 10 V", INPUT_RANGE_PM_10_V))
        Combo1.Items.Add(New VB6.ListBoxItem("+/- 20 V", INPUT_RANGE_PM_20_V))
        InputRangeIdChA = INPUT_RANGE_PM_50_MV


        Combo2.Items.Add(New VB6.ListBoxItem("+/- 40 mV", INPUT_RANGE_PM_40_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 50 mV", INPUT_RANGE_PM_50_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 80 mV", INPUT_RANGE_PM_80_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 100 mV", INPUT_RANGE_PM_100_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 200 mV", INPUT_RANGE_PM_200_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 400 mV", INPUT_RANGE_PM_400_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 500 mV", INPUT_RANGE_PM_500_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 800 mV", INPUT_RANGE_PM_800_MV))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 1 V", INPUT_RANGE_PM_1_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 2 V", INPUT_RANGE_PM_2_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 4 V", INPUT_RANGE_PM_4_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 5 V", INPUT_RANGE_PM_5_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 8 V", INPUT_RANGE_PM_8_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 10 V", INPUT_RANGE_PM_10_V))
        Combo2.Items.Add(New VB6.ListBoxItem("+/- 20 V", INPUT_RANGE_PM_20_V))
        InputRangeIdChB = INPUT_RANGE_PM_50_MV

        ' Set defaults
        Combo1.SelectedIndex = 1
        Combo2.SelectedIndex = 1



    End Sub

    'UPGRADE_WARNING: Event PrMF.Resize may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub PrMF_Resize(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Resize
        '    StatusL.ToolTipText = PrMF.Height
    End Sub

    'UPGRADE_NOTE: Form_Terminate was upgraded to Form_Terminate_Renamed. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    'UPGRADE_WARNING: PrMF event Form.Terminate has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6BA9B8D2-2A32-4B6E-8D36-44949974A5B4"'
    Private Sub Form_Terminate_Renamed()
        End
    End Sub

    Private Sub PrMF_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub


    ' Gets called every second, all DAQ commands get called from this
    Private Sub Hz_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Hz.Tick
        Dim zRateMin As Object
        Dim ixxx As Object
        'Static iCnt As Short
        Static zratesec As String
        'If iCnt = 0 Then
        '    Shape1.FillColor = &HFF&
        'End If
        'If iCnt = 1 Then
        '    Shape1.FillColor = &HFF00&
        '    iCnt = -1
        'End If
        'iCnt = iCnt + 1
        ' Tells VB to wait a short while before updating the interval timer
        If KeyHit = 1 Then
            KeyTimer = KeyTimer + 1
            If KeyTimer > 10 Then
                KeyHit = 0
                KeyTimer = 0
                IRate = Val(TimeT.Text)
                If IRate < GetNPrM() Then IRate = GetNPrM()
                If xRate > IRate * 60 Then xRate = IRate * 60
            End If
        End If

        'If Check6.value = 1 Then
        '    isubtrch1 = 1
        'Else
        '    isubtrch1 = 0
        'End If

        ' Parallel port
        'UPGRADE_WARNING: Couldn't resolve default property of object ixxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ixxx = Inp(Val("&H379"))
        'If OneTrueLiquid = 1 Then
        '    If (ixxx And 64) = 64 Then
        '        Shape1.FillColor = &HFF&
        '        IOk = 0
        '    Else
        '        Shape1.FillColor = &HFF00&
        '        IOk = 1
        '    End If
        'Else
        '    If (ixxx And 64) = 0 Then
        '        Shape1.FillColor = &HFF&
        '        IOk = 0
        '    Else
        '        Shape1.FillColor = &HFF00&
        '        IOk = 1
        '    End If
        'End If
        If (ixxx And 240) = 128 Then 'All 4 bits are low, note bit Busy is hardware inverted
            IOk = 0
            Shape1.FillColor = System.Drawing.ColorTranslator.FromOle(&HFF)
        Else
            IOk = 1
            Shape1.FillColor = System.Drawing.ColorTranslator.FromOle(&HFF00)
        End If

        'check liquid level
        If Check5.CheckState = 0 Then 'ignore liquid status not checked
            If (ixxx And 64) = 0 Then 'Ack low
                If Check4(0).Enabled = True Then
                    Check4(0).Enabled = False
                    Check4(0).BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                End If
            Else
                If Check4(0).Enabled = False Then
                    Check4(0).Enabled = True
                    Check4(0).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                End If
            End If

            If (ixxx And 128) = 128 Then 'Busy low
                If Check4(1).Enabled = True Then
                    Check4(1).Enabled = False
                    Check4(1).BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                End If
            Else
                If Check4(1).Enabled = False Then
                    Check4(1).Enabled = True
                    Check4(1).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                End If
            End If

            If (ixxx And 32) = 0 Then 'Paper Out low
                If Check4(2).Enabled = True Then
                    Check4(2).Enabled = False
                    Check4(2).BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                End If
                If Check4(3).Enabled = True Then
                    Check4(3).Enabled = False
                    Check4(3).BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                End If
            Else
                If Check4(2).Enabled = False Then
                    Check4(2).Enabled = True
                    Check4(2).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                End If
                If Check4(3).Enabled = False Then
                    Check4(3).Enabled = True
                    Check4(3).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                End If
            End If

            If (ixxx And 16) = 0 Then 'Select low
                If Check4(4).Enabled = True Then
                    Check4(4).Enabled = False
                    Check4(4).BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
                End If
            Else
                If Check4(4).Enabled = False Then
                    Check4(4).Enabled = True
                    Check4(4).BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
                End If
            End If
        End If

        List3.Items.Clear()
        List3.Items.Add("xRate = " & xRate)
        List3.Items.Add(" IRate = " & IRate)
        List3.Items.Add(" IRunning = " & iRunning)
        List3.Items.Add(" LiquidWait = " & LiquidWait)
        List3.Items.Add(" IOK = " & IOk)
        List3.Items.Add(" PassCount = " & PassCnt)
        'List3.AddItem " OneTrueLiquid = " & OneTrueLiquid
        List3.Items.Add(" ISets = " & ISets)
        List3.Items.Add(TimeOfDay & " " & Today)
        'List3.AddItem "check4(0) " & Check4(0)
        'List3.AddItem "check4(1) " & Check4(1)
        'List3.AddItem "check4(2) " & Check4(2)
        'List3.AddItem "check4(3) " & Check4(3)
        'List3.AddItem "check4(4) " & Check4(4)
        List3.Items.Add(IPrM & "/" & NPrM)
        'List3.AddItem "NPrM = " & GetNPrM()
        'UPGRADE_WARNING: Couldn't resolve default property of object ixxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        List3.Items.Add("ixxx = " & ixxx)
        List3.Items.Add("isubtrch1 = " & isubtrch1)
        If LiquidWait = 0 Then xRate = xRate - 1
        'UPGRADE_WARNING: Couldn't resolve default property of object zRateMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        zRateMin = Int(xRate / 60)
        zratesec = CStr(xRate Mod 60)
        If Len(zratesec) = 1 Then zratesec = "0" & zratesec
        'UPGRADE_WARNING: Couldn't resolve default property of object zRateMin. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        CntDwnL.Text = zRateMin & ":" & zratesec
        ' If xRate >= 0 code exits Hz Timer and no data is taken
        If xRate >= 0 Then Exit Sub
        If iRunning = 1 Then Exit Sub
        xRate = IRate * 60
        'Command1.Enabled = False
        iRunning = 1
        Do While (IPrM < NPrM)
            If Check4(IPrM).CheckState = 1 And Check4(IPrM).Enabled = True Then Exit Do
            IPrM = IPrM + 1
        Loop
        If Command1.Text = "DAQ Running - Press to Stop" And IPrM < NPrM Then
            StatusL.Text = "Taking Data"
            iRunning = 1
            RunFileL.BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
            RunFileL.ForeColor = System.Drawing.ColorTranslator.FromOle(0)
            '    Out Val("&H37a"), 15
            '    Out Val("&H37a"), 0
            FileOpen(3, xPath & "Run_num.ini", OpenMode.Input)
            Input(3, iiRun)
            iiRun = iiRun + 1
            iiFile = 1
            FileClose(3)
            FileOpen(3, xPath & "Run_num.ini", OpenMode.Output)
            PrintLine(3, iiRun)
            ' MsgBox xPath
            FileClose(3)
            '    RunNumL.Caption = Format(iiRun, "#0000") & "_" & Format(IPrM, "00")
            '    DataFileName = DataFilePath & "Run_" & Format(iiRun, "000000") & "_" & Format(IPrM, "00") & ".txt"
            '    RunFileL.Caption = DataFileName
            'Shape2(1).Visible = True
            _Shape2_1.Visible = True
            '    PulserWait.Interval = 30000
            '    PulserWait.Enabled = True
            '    Out Val("&H37a"), 15 'turn on pulser
            PassCnt = 0
            ' This is a function call, think TakeData()
            TakeData()
            '    Out Val("&H378"), IPrM + Data7On 'turn on pulser
            '    StatusL.Caption = "Pulser Turned ON, PrM = " & IPrM
            LiquidWait = 0
        End If

        If Command1.Text = "DAQ Running - Press to Stop" And IPrM = NPrM And (IOk = 1 Or Check5.CheckState = 1) Then
            RunFileL.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            RunFileL.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
            RunFileL.Text = "Data Not Taken yet No PrM is Selected"
            '    Open xPath & "Run_num.ini" For Input As #3
            '    Input #3, iiRun
            '    iiRun = iiRun + 1
            '    iiFile = 1
            '    Close #3
            '    Open xPath & "Run_num.ini" For Output As #3
            '    Print #3, iiRun
            '    Close #3
            '    RunNumL.Caption = Format(iiRun, "#0000") & "_" & Format(IPrM, "00")
            '    DataFileName = DataFilePath & "Run_" & Format(iiRun, "000000") & ".txt"
            '    RunFileL.Caption = DataFileName
            '    Open DataFileName For Append As #33
            '    Print #33, Date & "  " & Time
            '    Print #33, "Run Not Started Liquid not OK"
            '    Close #33
            LiquidWait = 1
            IPrM = 0
            OKWait.Enabled = True 'wait for liquid to be ok
            'Shape2(0).Visible = True
            _Shape2_0.Visible = True
            StatusL.Text = "Please select at least one PrM"

        End If

        If Command1.Text = "DAQ Running - Press to Stop" And IOk = 0 And Check5.CheckState = 0 Then
            RunFileL.BackColor = System.Drawing.ColorTranslator.FromOle(&HFF)
            RunFileL.ForeColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
            RunFileL.Text = "Data Not Taken yet Liquid not OK"
            '    Open xPath & "Run_num.ini" For Input As #3
            '    Input #3, iiRun
            '    iiRun = iiRun + 1
            '    iiFile = 1
            '    Close #3
            '    Open xPath & "Run_num.ini" For Output As #3
            '    Print #3, iiRun
            '    Close #3
            '    RunNumL.Caption = Format(iiRun, "#0000")
            '    DataFileName = DataFilePath & "Run_" & Format(iiRun, "000000") & ".txt"
            '    RunFileL.Caption = DataFileName
            '    Open DataFileName For Append As #33
            '    Print #33, Date & "  " & Time
            '    Print #33, "Run Not Started Liquid not OK"
            '    Close #33
            LiquidWait = 1
            OKWait.Enabled = True 'wait for liquid to be ok
            'Shape2(0).Visible = True
            _Shape2_0.Visible = True
            StatusL.Text = "Waiting for Liquid OK to take data"
        End If


        FileOpen(88, DataFileNameHist, OpenMode.Append)
        PrintLine(88, Today & " " & TimeOfDay & " Run Number = " & iiRun)
        If IOk = 1 Then
            PrintLine(88, "Liquid Status is OK")
        Else
            PrintLine(88, "Liquid Status is  Not OK")
        End If


        FileClose(88)


    End Sub

    ' Private Sub Label1_DblClick(Index As Integer)
    '    If Index = 1 Then
    '        If OneTrueLiquid = 1 Then
    '            OneTrueLiquid = 0
    '            Label1(1).ToolTipText = "High is good...Dbl Click to Change"
    '        Else
    '            OneTrueLiquid = 1
    '            Label1(1).ToolTipText = "Low is good...Dbl Click to Change"
    '        End If
    '    End If
    ' End Sub


    Private Sub Label2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Label2.Click
        Dim Index As Short = Label2.GetIndex(eventSender)

    End Sub

    Private Sub OKWait_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles OKWait.Tick
        If IOk = 0 Then Exit Sub
        Do While (IPrM < NPrM)
            If Check4(IPrM).CheckState = 1 And Check4(IPrM).Enabled = True Then Exit Do
            IPrM = IPrM + 1
        Loop
        If IPrM = NPrM Then
            IPrM = 0
            Exit Sub
        End If
        FileOpen(3, xPath & "Run_num.ini", OpenMode.Input)
        Input(3, iiRun)
        iiRun = iiRun + 1
        iiFile = 1
        FileClose(3)
        FileOpen(3, xPath & "Run_num.ini", OpenMode.Output)
        PrintLine(3, iiRun)
        FileClose(3)
        '    RunNumL.Caption = Format(iiRun, "#0000") & "_" & Format(IPrM, "00")
        '    DataFileName = DataFilePath & "Run_" & Format(iiRun, "000000") & "_" & Format(IPrM, "00") & ".txt"
        RunFileL.BackColor = System.Drawing.ColorTranslator.FromOle(&HFFFFFF)
        RunFileL.ForeColor = System.Drawing.ColorTranslator.FromOle(0)
        '    RunFileL.Caption = DataFileName
        '    PulserWait.Interval = 30000
        '    PulserWait.Enabled = True
        '    Out Val("&H37a"), 15 'turn on pulser
        PassCnt = 0
        TakeData()
        '    Out Val("&H378"), IPrM + Data7On 'turn on pulser
        StatusL.Text = "Pulser Turned ON, PrM = " & IPrM
        LiquidWait = 0
        OKWait.Enabled = False
        'Shape2(0).Visible = False
        _Shape2_0.Visible = False
        'Shape2(1).Visible = True
        _Shape2_1.Visible = True
        '    StatusL.Caption = "Taking Data"

    End Sub

    'UPGRADE_WARNING: Event PathT.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub PathT_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PathT.TextChanged
        LogPath = PathT.Text
    End Sub

    'UPGRADE_WARNING: Event IPathT.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub IPathT_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles IPathT.TextChanged
        ImagePath = IPathT.Text
    End Sub

    Private Sub PulserWait_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles PulserWait.Tick
        PulserWait.Interval = 15000
        PassCnt = PassCnt + 1
        '    Out Val("&H37a"), 0
        '    Out Val("&H37a"), 15
        'Out Val("&H37a"), 0
        If PassCnt <= ISets Then
            StatusL.Text = "Taking Data PrM = " & IPrM & " Pass = " & PassCnt
            ScopeWait.Enabled = True
            'Shape2(2).Visible = True
            _Shape2_2.Visible = True
            '       If PassCnt < ISets Then
            '               Out Val("&H378"), 8 + Data7On 'turn off pulse
            '               TakeData
            '       End If
        End If
        If PassCnt >= ISets + 1 Then
            IPrM = IPrM + 1
            Do While (IPrM < NPrM)
                If Check4(IPrM).CheckState = 1 And Check4(IPrM).Enabled = True Then Exit Do
                IPrM = IPrM + 1
            Loop
            If IPrM = NPrM Then
                '            Out Val("&H37a"), 0 'turn off pulser
                ' 64=2^6, turns off TPC veto for bit D6
                Out(Val("&H378"), 8 + Data7On + 64) 'turn off pulse
                PassCnt = 0
                PulserWait.Enabled = False
                'Shape2(1).Visible = False
                _Shape2_1.Visible = False
                iRunning = 0
                IPrM = 0
                'Command1.Enabled = True
                StatusL.Text = "Waiting for Next Interval"
                Exit Sub
            End If
            If IPrM < NPrM Then
                '            RunNumL.Caption = Format(iiRun, "#0000") & "_" & Format(IPrM, "00")
                '            DataFileName = DataFilePath & "Run_" & Format(iiRun, "000000") & "_" & Format(IPrM, "00") & ".txt"
                '            RunFileL.Caption = DataFileName
                '            PulserWait.Interval = 30000
                '            PulserWait.Enabled = True
                '           Out Val("&H37a"), 15 'turn on pulser
                PassCnt = 0
                TakeData()
                '            Out Val("&H378"), IPrM + Data7On 'turn on pulser
                'Shape2(1).Visible = True
                _Shape2_1.Visible = True
                StatusL.Text = "Pulser Turned ON, PrM = " & IPrM
                iiFile = 1
            End If
        End If
    End Sub

    Private Sub ScopeWait_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles ScopeWait.Tick
        Dim ii As Object
        Dim xsum As Object
        Dim kk As Object
        Dim xstring As Object
        'Dim TextStatus As Object
        Dim jj As Object
        Dim ccc As Object
        Dim zzzzzzz As Object
        'Dim abortAcquire As Object
        'Dim saveData As Object
        Dim sampleRateId As Object
        Dim AcquireData As Object
        Dim xxxx As Object
        Dim Max_Bins As Integer = -1
        ScopeWait.Enabled = False
        'Shape2(2).Visible = False
        _Shape2_2.Visible = False

        'read scope data here
        ' Check to see if the current device is really connected to something
        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        'UPGRADE_WARNING: Couldn't resolve default property of object zlocal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

        If (boardHandle = 0) Then
            ' Oops, the device isn't available, notify the user and exit
            MsgBox("BoardID " & boardID & " not available." & Chr(10) & "You need to select an available device" & Chr(10) & "for the Alazar boardID in the VB IDE.", MsgBoxStyle.OkOnly, "Startup Error")
            End
        End If








        ' New code for Alazar, takes data for channels A and B
        Dim preTriggerSamples As Integer
        Dim postTriggerSamples As Integer
        Dim samplesPerRecord As Integer
        Dim recordsPerCapture As Integer
        Dim captureTimeout_ms As Integer
        Dim channelMask As Integer
        Dim channelCount As Integer
        Dim channelsPerBoard As Integer
        Dim channelId As Integer
        Dim channel As Integer
        Dim bitsPerSample As Byte
        Dim bytesPerSample As Integer
        Dim bytesPerRecord As Integer
        Dim samplesPerBuffer As Integer
        Dim buffer() As Short
        Dim retCode As Integer
        Dim startTickCount As Integer
        Dim timeoutTickCount As Integer
        Dim captureDone As Boolean
        Dim captureTime_sec As Double
        Dim recordsPerSec As Double
        Dim bytesTransferred As Integer
        Dim record As Integer
        'Dim fileHandle As Short
        Dim transferTime_sec As Double
        Dim bytesPerSec As Double
        Dim sampleBitShift As Short
        Dim codeZero As Object
        Dim codeRange As Short
        Dim codeToValue As Double
        Dim scaleValueChA As Object
        Dim scaleValueChB As Double
        Dim offsetValue As Double
        ' These are variables needed for PrMF to run, some are redundant
        'UPGRADE_WARNING: Couldn't resolve default property of object xxxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xxxx = 1
        Dim arrWF() As Double 'array variable which will hold waveform values, Ben: this will have them after converted to volts, otherwise stored in buffer()
        Dim xinc As Double ' variable which will hold the x axis increment
        Dim trigpos As Integer ' variable which hold the timing trigger position
        Dim i As Integer ' counter variable
        Dim hUnits, vUnits As String ' variables for returning unit types
        'Dim strID As String 'tekvisa command
        Dim samplesPerSec As Double
        Dim xMarker As Integer

        FileOpen(33, DataFileName, OpenMode.Append)
        ' Print #33, Date & "  " & Time & "  Pass = " & PassCnt

        FileOpen(44, AllTracesFileName, OpenMode.Append)

        ' Set the default return code to indicate failure
        'UPGRADE_WARNING: Couldn't resolve default property of object AcquireData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AcquireData = False

        Dim DrawData As Boolean
        DrawData = True

        ' TODO: Select the number of pre-trigger samples per record
        preTriggerSamples = 500

        ' TODO: Select the number of post-trigger samples per record
        postTriggerSamples = 4508

        ' TODO: Select the number of records in the acquisition
        ' Ben: we would eventually like to average over these 10 captures
        recordsPerCapture = 10

        ' TODO: Select the amount of time, in milliseconds, to wait for the
        ' acquisiton to complete to on-board memory
        captureTimeout_ms = 100000

        ' TODO: Select which channels read from on-board memory (A, B, or both)
        channelMask = CHANNEL_A Or CHANNEL_B

        ' Select clock parameters as required
        ' 0, 2, 3 want 1 MSPS
        ' 1 and 4 want 5 MSPS
        ' Note on 4/10/2014, 0 is temporarily a short purity monitor
        ' Note on 4/15/2014, the 5 MSPS rate doesn't work for 0, I don't know why
        If IPrM = 0 Or IPrM = 2 Or IPrM = 3 Then
            'If IPrM = 2 Or IPrM = 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object sampleRateId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sampleRateId = SAMPLE_RATE_2MSPS
            samplesPerSec = 2000000
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object sampleRateId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sampleRateId = SAMPLE_RATE_5MSPS
            samplesPerSec = 5000000
        End If

        ' Find the number of enabled channels from the channel mask
        channelsPerBoard = 2
        channelCount = 0
        For channel = 0 To channelsPerBoard - 1
            channelId = 2 ^ channel
            If (channelMask And channelId) = channelId Then
                channelCount = channelCount + 1
            End If
        Next channel

        If ((channelCount < 1) Or (channelCount > channelsPerBoard)) Then
            MsgBox("Error: Invalid channel mask " & Hex(channelMask))
            Exit Sub
        End If

        ' Calculate the size of each record in bytes
        samplesPerRecord = preTriggerSamples + postTriggerSamples

        bitsPerSample = 12
        bytesPerSample = (bitsPerSample + 7) \ 8
        bytesPerRecord = bytesPerSample * samplesPerRecord

        ' Allocate an array of 16-bit signed integers to receive one full record
        ' Note that the buffer must be at least 16 samples larger than the number of samples to transfer
        samplesPerBuffer = samplesPerRecord + 16
        ReDim buffer(samplesPerBuffer - 1)

        ' Create a data file if required
        'If (saveData) Then
        'fileHandle = FreeFile()
        'FileOpen(fileHandle, "file.txt", OpenMode.Output)
        'End If

        ' Configure the number of samples per record
        retCode = AlazarSetRecordSize(boardHandle, preTriggerSamples, postTriggerSamples)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetRecordSize failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Configure the number of records in the acquisition
        retCode = AlazarSetRecordCount(boardHandle, recordsPerCapture)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetRecordCount failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Update status
        ' TextStatus = TextStatus & "Capturing " & recordsPerCapture & " records" & vbCrLf
        ' TextStatus.Refresh

        ' Arm the board to begin the acquisition
        retCode = AlazarStartCapture(boardHandle)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarStartCapture failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Wait for the board to capture all records to on-board memory
        startTickCount = GetTickCount()
        timeoutTickCount = startTickCount + captureTimeout_ms
        captureDone = False

        Do While Not captureDone
            If AlazarBusy(boardHandle) = 0 Then
                ' All records have been captured to on-board memory
                captureDone = True
            ElseIf (GetTickCount() > timeoutTickCount) Then
                ' The capture timeout expired before the capture completed
                ' The board may not be triggering, or the capture timeout is too short.
                MsgBox("Error: Capture timeout after " & captureTimeout_ms & "ms -- verify trigger")
                Exit Do
                ' ElseIf (abortAcquire) Then
                ' The Abort button was pressed
                'Exit Do
            Else
                ' The capture is in progress
                If GetInputState <> 0 Then System.Windows.Forms.Application.DoEvents()
                Sleep(10)
            End If
        Loop

        ' Abort the acquisition and exit if the acquisition did not complete
        If Not captureDone Then
            retCode = AlazarAbortCapture(boardHandle)
            If (retCode <> ApiSuccess) Then
                MsgBox("Error: AlazarAbortCapture " & AlazarErrorToText(retCode))
            End If
            Exit Sub
        End If

        ' Report time to capture all records to on-board memory
        captureTime_sec = (GetTickCount() - startTickCount) / 1000
        If (captureTime_sec > 0) Then
            recordsPerSec = recordsPerCapture / captureTime_sec
        Else
            recordsPerSec = 0
        End If
        ' TextStatus = TextStatus & "Captured " & recordsPerCapture & " records in " & _
        ''    captureTime_sec & " sec (" & Format(recordsPerSec, "0.00e+") & " records / sec)" & vbCrLf
        ' TextStatus.Refresh

        ' Calculate scale factors to convert sample values to volts:

        ' (a) This 14-bit sample code represents a 0V input
        'UPGRADE_WARNING: Couldn't resolve default property of object codeZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        codeZero = 2 ^ (bitsPerSample - 1) - 0.5

        ' (b) This is the range of 14-bit sample codes with respect to 0V level
        codeRange = 2 ^ (bitsPerSample - 1) - 0.5

        ' (c) 14-bit sample codes are stored in the most signficant bits of each 16-bit sample value
        sampleBitShift = 8 * bytesPerSample - bitsPerSample

        ' (d) Mutiply a 14-bit sample code by this amount to get a 16-bit sample value
        codeToValue = 2 ^ sampleBitShift

        ' (e) Subtract this amount from a 16-bit sample value to remove the 0V offset
        'UPGRADE_WARNING: Couldn't resolve default property of object codeZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        offsetValue = codeZero * codeToValue

        ' (f) Multiply a 16-bit sample value by this factor to convert it to volts
        'UPGRADE_WARNING: Couldn't resolve default property of object scaleValueChA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        scaleValueChA = InputRangeIdToVolts(InputRangeIdChA) / (codeRange * codeToValue)
        scaleValueChB = InputRangeIdToVolts(InputRangeIdChB) / (codeRange * codeToValue)

        ' Transfer records from on-board memory to our buffer, one record at a time
        ' TextStatus = TextStatus & "Transferring " & recordsPerCapture & " records" & vbCrLf
        ' TextStatus.Refresh

        startTickCount = GetTickCount()
        bytesTransferred = 0
        'UPGRADE_ISSUE: PictureBox property PictureBox.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'PictureBox.AutoRedraw = True
        'UPGRADE_ISSUE: PictureBox property Picture1.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Picture1.AutoRedraw = True
        Dim channelName As String
        Dim scaleValue As Double
        Dim sampleInRecord As Integer
        Dim sampleValue As Integer
        Dim sampleVolts As Double
        'Dim Timestamp_samples As Decimal
        'Dim timestamp_address As Integer
        'Dim timestamp_highpart As Integer
        'Dim timestamp_lowpart As Integer
        For record = 0 To recordsPerCapture - 1

            ' Erase the previous record
            If DrawData = True Then
                'UPGRADE_ISSUE: PictureBox method PictureBox.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'PictureBox.Cls()
                PictureBox = Nothing
                'UPGRADE_ISSUE: PictureBox method Picture1.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'Picture1.Cls()
                PictureBox = Nothing
            End If

            For channel = 0 To channelsPerBoard - 1

                ' Get channel Id from channel index
                channelId = 2 ^ channel

                ' Skip this channel unless it is in channel mask
                If (channelMask And channelId) <> 0 Then

                    ' Erase the contents of the arrWF array
                    ReDim arrWF(1)

                    ' Transfer one full record from on-board memory to our buffer
                    retCode = AlazarRead(boardHandle, channelId, buffer(0), bytesPerSample, record + 1, -preTriggerSamples, samplesPerRecord)
                    If (retCode <> ApiSuccess) Then
                        MsgBox("Error: AlazarRead failed -- " & AlazarErrorToText(retCode))
                        Exit Sub
                    End If

                    ' TODO: Process record here.
                    '
                    ' Samples values are arranged contiguously in the buffer, with
                    ' 12-bit sample codes in the most significant bits of each 16-bit
                    ' sample value.
                    '
                    ' Sample codes are unsigned by default so that:
                    ' - a sample code of 0x000 represents a negative full scale input signal;
                    ' - a sample code of 0x800 represents a ~0V signal;
                    ' - a sample code of 0xFFF represents a positive full scale input signal.

                    ' Save record to file
                    ' If saveData = True Then

                    ' Find name and scale factor of this channel
                    Select Case channelId
                        Case CHANNEL_A
                            channelName = "A"
                            'UPGRADE_WARNING: Couldn't resolve default property of object scaleValueChA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            scaleValue = scaleValueChA
                        Case CHANNEL_B
                            channelName = "B"
                            scaleValue = scaleValueChB
                        Case Else
                            MsgBox("Error: Invalid channelId " & channelId)
                            Exit Sub
                    End Select

                    ' Delimit the start of this record in the file
                    '                    Print #fileHandle, "--> CH" & channelName & " record " & record + 1 & " begin"
                    '                    Print #fileHandle, ""

                    ' Write record samples to file
                    For sampleInRecord = 0 To samplesPerRecord - 1

                        ' Get a sample value from the buffer
                        ' Note that the digitizer returns 16-bit unsigned sample values
                        ' that are stored in a 16-bit signed integer array, so convert
                        ' a signed 16-bit value to unsigned.
                        If (buffer(sampleInRecord) < 0) Then
                            sampleValue = buffer(sampleInRecord) + 65536
                        Else
                            sampleValue = buffer(sampleInRecord)
                        End If

                        ' Convert the sample value to volts
                        sampleVolts = scaleValue * (sampleValue - offsetValue)

                        ' Store new sampleVolts
                        ReDim Preserve arrWF(UBound(arrWF) + 1)
                        arrWF(sampleInRecord) = sampleVolts

                        ' Write the sample value in volts to file
                        ' Print #33, Str(sampleInRecord) & vbTab & sampleVolts
                        ' This is what I think displays units
                        ' sampleInRecord runs from 0 to 1023
                        ' Sampling rate is 5000000 Hz
                        ' To get seconds: sampleInRecord/5000000
                        ' Print #33, Str(sampleInRecord / 5000000) & " sec." & vbTab & Format(sampleVolts, "0.0000") & " V"
                        ' Print #33, Str(sampleInRecord) & vbTab & Format(sampleVolts, "0.0000")

                        PrintLine(44, Str(sampleInRecord / samplesPerSec) & " sec." & vbTab & VB6.Format(sampleVolts, "0.000000") & " V")


                    Next sampleInRecord

                    'retCode = AlazarGetTriggerTimestamp(boardHandle, record, timestamp_samples)
                    ' retCode = AlazarGetTriggerAddress(boardHandle, record, timestamp_address, timestamp_highpart, timestamp_lowpart)
                    'If (retCode <> ApiSuccess) Then
                    '    MsgBox ("Error: AlazarGetTriggerAddress failed -- " & AlazarErrorToText(retCode))
                    '    Exit Sub
                    'End If

                    'trigpos = timestamp_samples
                    'trigpos = timestamp_highpart
                    'trigpos = 128 * 2
                    trigpos = preTriggerSamples

                    xinc = 1.0# / samplesPerSec ' Length of time for each array element
                    vUnits = CStr(1)
                    hUnits = CStr(1)
                    xIncr = xinc
                    XTrig = trigpos ' time when trigger fired
                    xMarker = XTrig
                    If IsArray(arrWF) Then ' check to be sure returned value is an array
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        zzzzzzz = UBound(arrWF) - LBound(arrWF) + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ToolTip1.SetToolTip(StatusL, "Scope Array Size = " & zzzzzzz)
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If zzzzzzz > 100000 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            MsgBox("Array too Large  size= " & zzzzzzz, , "Fix settings on scope")
                        End If
                        IDPoints = UBound(arrWF) - LBound(arrWF) + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object ccc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ccc = 0
                    Else
                        MsgBox("Error not array " & channel)
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    jj = 0
                    ' Print #33, "Ch" & (channel + 1) & " " & xInc & " " & hUnits & " " & XTrig & " " & vUnits
                    For i = LBound(arrWF) To UBound(arrWF)
                        'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ' Histo(channel + 1, jj) = arrWF(i) * 1000000.0#
                        ' Print #33, arrWF(i)
                        ' xData(channel, jj) = arrWF(i) - xDataRef(channel, jj)
                        'xData(channel, jj) = xData(channel, jj) + arrWF(i)
                        ' xData(channel, jj) = arrWF(i)
                        If record = 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            xData(channel, jj) = arrWF(i)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            xData(channel, jj) = xData(channel, jj) + arrWF(i)
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        jj = jj + 1
                    Next i
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Max_Bins = jj






                    '                    ' Delimit the end of this record in the file
                    '                    Print #fileHandle, "<-- CH " & channelName & " record " & record + 1 & " end"
                    '                    Print #fileHandle, ""

                    ' End If  ' saveData

                    ' Draw record on screen
                    If DrawData = True Then
                        Call DrawRecord(buffer, samplesPerRecord, channelId)
                    End If

                    bytesTransferred = bytesTransferred + bytesPerRecord

                End If ' (channelMask And channelId) <> 0

            Next channel

            ' PictureBox.AutoRedraw = True
            ' Save picture
            ' SavePicture PictureBox.Image, "C:\Example.bmp"
            'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            PictureBox.Image.Save(ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".raw.bmp")
            ' Convert from bmp to jpeg
            ' BmpToJpeg "C:\Example" & ".bmp", "C:\Example" & ".jpg", 100
            BmpToJpeg(ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".raw.bmp", ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".raw.jpg", 100)





            ' If the abort button was pressed, then stop reading records

            If GetInputState <> 0 Then System.Windows.Forms.Application.DoEvents()
            'UPGRADE_WARNING: Couldn't resolve default property of object abortAcquire. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'If abortAcquire Then
            'UPGRADE_WARNING: Couldn't resolve default property of object TextStatus. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'TextStatus = TextStatus + "Aborted" & vbCrLf
            'UPGRADE_WARNING: Couldn't resolve default property of object TextStatus.Refresh. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'TextStatus.Refresh()
            'Exit For
            'End If

        Next record



        ' Average out the signal samples
        For channel = 0 To channelsPerBoard - 1
            ' For i = LBound(xData, 2) To UBound(xData, 2)
            For i = 0 To preTriggerSamples + postTriggerSamples - 1
                xData(channel, i) = xData(channel, i) / recordsPerCapture
                PrintLine(33, Str(i / samplesPerSec) & " sec." & vbTab & VB6.Format(xData(channel, i), "0.000000") & " V")
                ' xData(channel, i) = xData(channel, i) - xDataRef(channel, i)
            Next i
        Next channel


        ' Close the data file
        'If (saveData) Then
        '    Close #fileHandle
        'End If




        ' Print #33, "Channel Count = " & ChCnt & " Points = " & Max_Bins & " " & xInc & " " & hUnits & " " & XTrig & " " & vUnits
        For i = 0 To Max_Bins - 1
            'UPGRADE_WARNING: Couldn't resolve default property of object xstring. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xstring = Str(-XTrig * xinc + i * xinc)
            For kk = 0 To 1
                'UPGRADE_WARNING: Couldn't resolve default property of object kk. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xstring. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If Check1(kk).CheckState = 1 Then xstring = xstring & "," & Str(xData(kk, i))
            Next kk
            ' Print #33, xstring
        Next i
        ' Close #33

        ' FileCopy DataFileName, ImagePath & "Run_" & Format(iiRun, "000000") & ".txt"
        ' Display results
        transferTime_sec = (GetTickCount() - startTickCount) / 1000
        If (transferTime_sec > 0) Then
            bytesPerSec = bytesTransferred / transferTime_sec
        Else
            bytesPerSec = 0
        End If

        '    TextStatus = TextStatus & _
        ''        "Transferred " & Format(bytesTransferred, "0.00e+") & " bytes in " & transferTime_sec & _
        ''        " sec (" & Format(bytesPerSec, "0.00e+") & " bytes per sec)" & vbCrLf
        '    TextStatus.Refresh

        ' Set the return code to indicate success
        'UPGRADE_WARNING: Couldn't resolve default property of object AcquireData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AcquireData = True

        ' End of new code needed for Alazar digitizer

























        '    xxxx = 1
        '  ' declare variables
        '   Dim arrWF As Variant 'array variable which will hold waveform values
        '   Dim xInc As Double ' variable which will hold the x axis increment
        '   Dim trigpos As Long ' variable which hold the timing trigger position
        '   Dim i As Long  ' counter variable
        '   Dim hUnits As String, vUnits As String   ' variables for returning unit types
        '   Dim strID As String 'tekvisa command

        ''   On Error GoTo cmdGetWFMErr1
        '   On Error GoTo 0
        '   On Error Resume Next
        'Open DataFileName For Append As #33
        'Print #33, Date & "  " & Time & "  Pass = " & PassCnt

        '   'CH1 is the OCX built-in constant specifying Channel 1
        'For kk = 0 To 3
        '    If Check1(kk).value = 1 Then
        '        If kk = 0 Then
        '            Tvc1.WriteString "MATH1:DEFIne ""CH1-REF1""" ' remove noise, performing subtraction, arrWF is an array that stores the waveform, the equivalent to GetWaveForm might be AlazarRead
        '            Call Tvc1.GetWaveform(MATH1, arrWF, xInc, trigpos, vUnits, hUnits)
        ''            Call Tvc1.GetWaveform(CH1, arrWF, xInc, trigpos, vUnits, hUnits)
        '        End If
        '        If kk = 1 Then
        '            Tvc1.WriteString "MATH2:DEFIne ""CH2-REF2""" ' remove noise
        '            Call Tvc1.GetWaveform(MATH2, arrWF, xInc, trigpos, vUnits, hUnits)
        ''            Call Tvc1.GetWaveform(CH2, arrWF, xInc, trigpos, vUnits, hUnits)
        '        End If
        '        If kk = 2 Then Call Tvc1.GetWaveform(CH3, arrWF, xInc, trigpos, vUnits, hUnits)
        '        If kk = 3 Then Call Tvc1.GetWaveform(CH4, arrWF, xInc, trigpos, vUnits, hUnits)
        '        Label3.Caption = xInc & hUnits
        '        Label4.Caption = vUnits
        '        xIncr = xInc
        '        XTrig = trigpos
        '        xMarker = XTrig
        '        If IsArray(arrWF) Then  ' check to be sure returned value is an array
        '            zzzzzzz = UBound(arrWF) - LBound(arrWF) + 1
        '            StatusL.ToolTipText = "Scope Array Size = " & zzzzzzz
        '            If zzzzzzz > 10000 Then
        '                MsgBox "Array too Large  size= " & zzzzzzz, , "Fix settings on scope"
        '            End If
        '            IDPoints = UBound(arrWF) - LBound(arrWF) + 1
        '            ccc = 0
        '        Else
        '            MsgBox "Error not array " & kk
        '        End If
        '        jj = 0
        '        Print #33, "Ch" & (kk + 1) & " " & xInc & " " & hUnits & " " & XTrig & " " & vUnits
        '        For i = LBound(arrWF) To UBound(arrWF)
        '            Histo(kk + 1, jj) = arrWF(i) * 1000000#
        '            'Print #33, arrWF(i)
        '            xData(kk, jj) = arrWF(i)
        '            jj = jj + 1
        '        Next i
        '        Max_Bins = jj
        '    End If
        'Next kk
        'Print #33, "Channel Count = " & ChCnt & " Points = " & Max_Bins & " " & xInc & " " & hUnits & " " & XTrig & " " & vUnits
        'For i = 0 To Max_Bins - 1
        '    xstring = Str(-XTrig * xInc + i * xInc)
        '    For kk = 0 To 3
        '        If Check1(kk).value = 1 Then xstring = xstring & "," & Str(xData(kk, i))
        '    Next kk
        '    Print #33, xstring
        'Next i
        'Close #33


        For i = 0 To 1 'now create smoothed data
            'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xsum = 0
            For ii = 0 To ISmooth - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xsum = xsum + xData(i, ii)
            Next ii
            For ii = 0 To ISmooth / 2 - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xData(i + 4, ii) = xsum / ISmooth
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
            Next ii
            'xData(i + 4) = xsum / ISmooth

            For ii = ISmooth / 2 To Max_Bins - ISmooth / 2 - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xsum = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For jj = ii - ISmooth / 2 To ii + ISmooth / 2
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xsum = xsum + xData(i, jj)
                Next jj
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xData(i + 4, ii) = xsum / (ISmooth + 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
                '        xsum = xsum - xData(i, ii - ISmooth) + xData(i, ii)
                '        xData(i + 4, ii - ISmooth) = xsum / ISmooth
                '        Histo(i + 5, ii - ISmooth) = 1000000# * xsum / ISmooth
            Next ii
            For ii = Max_Bins - ISmooth / 2 To Max_Bins
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xData(i + 4, ii) = xsum / ISmooth
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                ' Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
            Next ii
        Next i



        Call DrawAveragedSmoothedRecord(xData, preTriggerSamples + postTriggerSamples, channelsPerBoard)



        ' Print out the noise samples, smoothed
        For channel = 0 To channelsPerBoard - 1
            ' For i = LBound(xData, 2) To UBound(xData, 2)
            For i = 0 To preTriggerSamples + postTriggerSamples - 1
                PrintLine(33, Str(i / samplesPerSec) & " sec." & vbTab & VB6.Format(xDataRef(channel + 4, i), "0.000000") & " V")
                ' xData(channel, i) = xData(channel, i) - xDataRef(channel, i)
            Next i
        Next channel

        ' Print out the signal samples
        For channel = 0 To channelsPerBoard - 1
            ' For i = LBound(xData, 2) To UBound(xData, 2)
            For i = 0 To preTriggerSamples + postTriggerSamples - 1
                PrintLine(33, Str(i / samplesPerSec) & " sec." & vbTab & VB6.Format(xData(channel + 4, i), "0.000000") & " V")
                ' xData(channel, i) = xData(channel, i) - xDataRef(channel, i)
            Next i
        Next channel


        ' Close the data file
        'If (saveData) Then
        'FileClose(fileHandle)
        'End If

        FileClose(44)
        FileClose(33)
        FileCopy(DataFileName, ImagePath & "Run_" & VB6.Format(iiRun, "000000") & ".txt")





        Call Anal_Data()
        'If Hist.Visible = True Then Hist.Data_Ready.Text = "1"
        'If PassCnt = 3 Then Out Val("&H37a"), 0 'turn off pulser
        '' 64=2^6, turns off TPC veto for bit D6, we take TPC data again, out of date 1/23/2013
        If PassCnt = 3 Then Out(Val("&H378"), 8 + Data7On) 'turn off pulser

        Exit Sub
        'rudimentary error trapping
cmdGetWFMErr1:
        'If PassCnt = ISets Then Out Val("&H37a"), 0 'turn off pulser
        If PassCnt = ISets Then Out(Val("&H378"), 8 + Data7On) 'turn off pulser
        FileClose(33)
        MsgBox("Error: " & Err.Number & ", " & Err.Description)

    End Sub

    'UPGRADE_WARNING: Event Text2.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Text2_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text2.TextChanged
        ISmooth = Val(Text2.Text)
        If ISmooth Mod 2 <> 0 Then ISmooth = ISmooth + 1
    End Sub

    'UPGRADE_WARNING: Event Text3.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Text3_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text3.TextChanged
        RMSCut = Val(Text3.Text)
        If RMSCut < 3 Then RMSCut = 3.0#
    End Sub

    'UPGRADE_WARNING: Event Text4.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub Text4_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text4.TextChanged
        Static ilocal As Short
        If ilocal = 1 Then Exit Sub
        ISets = Val(Text4.Text)
        If ISets > 5 Then
            ISets = 5
            ilocal = 1
            Text4.Text = CStr(ISets)
        End If
        If ISets < 1 Then
            ISets = 1
            ilocal = 1
            Text4.Text = CStr(ISets)
        End If
        ilocal = 0

    End Sub

    Private Sub HVWait_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles HVWait.Tick
        Dim ii As Object
        Dim xsumref As Object
        'Dim TextStatus As Object
        Dim jj As Object
        Dim ccc As Object
        Dim zzzzzzz As Object
        'Dim abortAcquire As Object
        'Dim saveData As Object
        Dim sampleRateId As Object
        Dim AcquireData As Object
        Dim xxxx As Object
        ' Function accumulates noise for later correction
        Dim Max_Bins As Integer = -1
        Dim xMarker As Integer

        HVWait.Enabled = False




        ' Check to see if the current device is really connected to something
        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        'UPGRADE_WARNING: Couldn't resolve default property of object zlocal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (boardHandle = 0) Then
            ' Oops, the device isn't available, notify the user and exit
            MsgBox("BoardID " & boardID & " not available." & Chr(10) & "You need to select an available device" & Chr(10) & "for the Alazar boardID in the VB IDE.", MsgBoxStyle.OkOnly, "Startup Error")
            End
        End If





        ' New code for Alazar, takes data for channels A and B
        Dim preTriggerSamples As Integer
        Dim postTriggerSamples As Integer
        Dim samplesPerRecord As Integer
        Dim recordsPerCapture As Integer
        Dim captureTimeout_ms As Integer
        Dim channelMask As Integer
        Dim channelCount As Integer
        Dim channelsPerBoard As Integer
        Dim channelId As Integer
        Dim channel As Integer
        Dim bitsPerSample As Byte
        Dim bytesPerSample As Integer
        Dim bytesPerRecord As Integer
        Dim samplesPerBuffer As Integer
        Dim buffer() As Short
        Dim retCode As Integer
        Dim startTickCount As Integer
        Dim timeoutTickCount As Integer
        Dim captureDone As Boolean
        Dim captureTime_sec As Double
        Dim recordsPerSec As Double
        Dim bytesTransferred As Integer
        Dim record As Integer
        'Dim fileHandle As Short
        'Dim transferTime_sec As Double
        'Dim bytesPerSec As Double
        Dim sampleBitShift As Short
        Dim codeZero As Object
        Dim codeRange As Short
        Dim codeToValue As Double
        Dim scaleValueChA As Object
        Dim scaleValueChB As Double
        Dim offsetValue As Double
        ' These are variables needed for PrMF to run, some are redundant
        'UPGRADE_WARNING: Couldn't resolve default property of object xxxx. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xxxx = 1
        Dim arrWF() As Double 'array variable which will hold waveform values, Ben: this will have them after converted to volts, otherwise stored in buffer()
        Dim xinc As Double ' variable which will hold the x axis increment
        Dim trigpos As Integer ' variable which hold the timing trigger position
        Dim i As Integer ' counter variable
        Dim hUnits, vUnits As String ' variables for returning unit types
        'Dim strID As String 'tekvisa command
        Dim samplesPerSec As Double

        FileClose(33)
        FileOpen(33, DataFileName, OpenMode.Append)
        PrintLine(33, Today & "  " & TimeOfDay & "  Pass = " & PassCnt)

        FileClose(44)
        FileOpen(44, AllTracesFileName, OpenMode.Append)
        PrintLine(44, Today & "  " & TimeOfDay & "  Pass = " & PassCnt)


        ' Set the default return code to indicate failure
        'UPGRADE_WARNING: Couldn't resolve default property of object AcquireData. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        AcquireData = False

        Dim DrawData As Boolean
        DrawData = True

        ' TODO: Select the number of pre-trigger samples per record
        ' preTriggerSamples = 128

        ' TODO: Select the number of post-trigger samples per record
        ' postTriggerSamples = 896

        preTriggerSamples = 500

        ' TODO: Select the number of post-trigger samples per record
        postTriggerSamples = 4508

        ' TODO: Select the number of records in the acquisition
        ' Ben: we would eventually like to average over 10 captures
        recordsPerCapture = 10

        ' TODO: Select the amount of time, in milliseconds, to wait for the
        ' acquisiton to complete to on-board memory
        captureTimeout_ms = 100000

        ' TODO: Select which channels read from on-board memory (A, B, or both)
        channelMask = CHANNEL_A Or CHANNEL_B

        ' Select clock parameters as required
        ' 0, 2, 3 want 1 MSPS
        ' 1 and 4 want 5 MSPS
        ' Note on 4/10/2014, 0 is temporarily short
        ' Note on 4/15/2014, the 5 MSPS rate doesn't work for 0, I don't know why
        If IPrM = 0 Or IPrM = 2 Or IPrM = 3 Then
            'If IPrM = 2 Or IPrM = 3 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object sampleRateId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sampleRateId = SAMPLE_RATE_2MSPS
            samplesPerSec = 2000000
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object sampleRateId. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            sampleRateId = SAMPLE_RATE_5MSPS
            samplesPerSec = 5000000
        End If

        ' Find the number of enabled channels from the channel mask
        channelsPerBoard = 2
        channelCount = 0
        For channel = 0 To channelsPerBoard - 1
            channelId = 2 ^ channel
            If (channelMask And channelId) = channelId Then
                channelCount = channelCount + 1
            End If
        Next channel

        If ((channelCount < 1) Or (channelCount > channelsPerBoard)) Then
            MsgBox("Error: Invalid channel mask " & Hex(channelMask))
            Exit Sub
        End If

        ' Calculate the size of each record in bytes
        samplesPerRecord = preTriggerSamples + postTriggerSamples

        bitsPerSample = 12
        bytesPerSample = (bitsPerSample + 7) \ 8
        bytesPerRecord = bytesPerSample * samplesPerRecord

        ' Allocate an array of 16-bit signed integers to receive one full record
        ' Note that the buffer must be at least 16 samples larger than the number of samples to transfer
        samplesPerBuffer = samplesPerRecord + 16
        ReDim buffer(samplesPerBuffer - 1)

        ' Create a data file if required
        'If (saveData) Then
        ' fileHandle = FreeFile()
        ' FileOpen(fileHandle, "file.txt", OpenMode.Output)
        ' End If

        ' Configure the number of samples per record
        retCode = AlazarSetRecordSize(boardHandle, preTriggerSamples, postTriggerSamples)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetRecordSize failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Configure the number of records in the acquisition
        retCode = AlazarSetRecordCount(boardHandle, recordsPerCapture)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetRecordCount failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Update status
        ' TextStatus = TextStatus & "Capturing " & recordsPerCapture & " records" & vbCrLf
        ' TextStatus.Refresh

        ' Arm the board to begin the acquisition
        retCode = AlazarStartCapture(boardHandle)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarStartCapture failed -- " & AlazarErrorToText(retCode))
            Exit Sub
        End If

        ' Wait for the board to capture all records to on-board memory
        startTickCount = GetTickCount()
        timeoutTickCount = startTickCount + captureTimeout_ms
        captureDone = False

        Do While Not captureDone
            If AlazarBusy(boardHandle) = 0 Then
                ' All records have been captured to on-board memory
                captureDone = True
            ElseIf (GetTickCount() > timeoutTickCount) Then
                ' The capture timeout expired before the capture completed
                ' The board may not be triggering, or the capture timeout is too short.
                MsgBox("Error: Capture timeout after " & captureTimeout_ms & "ms -- verify trigger")
                Exit Do
            Else
                ' The capture is in progress
                If GetInputState <> 0 Then System.Windows.Forms.Application.DoEvents()
                Sleep(10)
            End If
        Loop

        ' Abort the acquisition and exit if the acquisition did not complete
        If Not captureDone Then
            retCode = AlazarAbortCapture(boardHandle)
            If (retCode <> ApiSuccess) Then
                MsgBox("Error: AlazarAbortCapture " & AlazarErrorToText(retCode))
            End If
            Exit Sub
        End If

        ' Report time to capture all records to on-board memory
        captureTime_sec = (GetTickCount() - startTickCount) / 1000
        If (captureTime_sec > 0) Then
            recordsPerSec = recordsPerCapture / captureTime_sec
        Else
            recordsPerSec = 0
        End If
        ' TextStatus = TextStatus & "Captured " & recordsPerCapture & " records in " & _
        ''    captureTime_sec & " sec (" & Format(recordsPerSec, "0.00e+") & " records / sec)" & vbCrLf
        ' TextStatus.Refresh

        ' Calculate scale factors to convert sample values to volts:

        ' (a) This 12-bit sample code represents a 0V input
        'UPGRADE_WARNING: Couldn't resolve default property of object codeZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        codeZero = 2 ^ (bitsPerSample - 1) - 0.5

        ' (b) This is the range of 14-bit sample codes with respect to 0V level
        codeRange = 2 ^ (bitsPerSample - 1) - 0.5

        ' (c) 12-bit sample codes are stored in the most signficant bits of each 16-bit sample value
        sampleBitShift = 8 * bytesPerSample - bitsPerSample

        ' (d) Mutiply a 12-bit sample code by this amount to get a 16-bit sample value
        codeToValue = 2 ^ sampleBitShift

        ' (e) Subtract this amount from a 12-bit sample value to remove the 0V offset
        'UPGRADE_WARNING: Couldn't resolve default property of object codeZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        offsetValue = codeZero * codeToValue

        ' (f) Multiply a 16-bit sample value by this factor to convert it to volts
        'UPGRADE_WARNING: Couldn't resolve default property of object scaleValueChA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        scaleValueChA = InputRangeIdToVolts(InputRangeIdChA) / (codeRange * codeToValue)
        scaleValueChB = InputRangeIdToVolts(InputRangeIdChB) / (codeRange * codeToValue)

        ' Transfer records from on-board memory to our buffer, one record at a time
        ' TextStatus = TextStatus & "Transferring " & recordsPerCapture & " records" & vbCrLf
        ' TextStatus.Refresh

        startTickCount = GetTickCount()
        bytesTransferred = 0
        'UPGRADE_ISSUE: PictureBox property PictureBox.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'PictureBox.AutoRedraw = True
        'UPGRADE_ISSUE: PictureBox property Picture1.AutoRedraw was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Picture1.AutoRedraw = True
        Dim channelName As String
        Dim scaleValue As Double
        Dim sampleInRecord As Integer
        Dim sampleValue As Integer
        Dim sampleVolts As Double
        For record = 0 To recordsPerCapture - 1

            ' Erase the previous record
            If DrawData = True Then
                'UPGRADE_ISSUE: PictureBox method PictureBox.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'PictureBox.Cls()
                PictureBox = Nothing
                'UPGRADE_ISSUE: PictureBox method Picture1.Cls was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'Picture1.Cls()
                Picture1 = Nothing
            End If

            For channel = 0 To channelsPerBoard - 1

                ' Get channel Id from channel index
                channelId = 2 ^ channel

                ' Skip this channel unless it is in channel mask
                If (channelMask And channelId) <> 0 Then

                    ' Erase the contents of the arrWF array
                    ReDim arrWF(1)

                    ' Transfer one full record from on-board memory to our buffer
                    retCode = AlazarRead(boardHandle, channelId, buffer(0), bytesPerSample, record + 1, -preTriggerSamples, samplesPerRecord)
                    If (retCode <> ApiSuccess) Then
                        MsgBox("Error: AlazarRead failed -- " & AlazarErrorToText(retCode))
                        Exit Sub
                    End If

                    ' TODO: Process record here.
                    '
                    ' Samples values are arranged contiguously in the buffer, with
                    ' 12-bit sample codes in the most significant bits of each 16-bit
                    ' sample value.
                    '
                    ' Sample codes are unsigned by default so that:
                    ' - a sample code of 0x000 represents a negative full scale input signal;
                    ' - a sample code of 0x800 represents a ~0V signal;
                    ' - a sample code of 0xFFF represents a positive full scale input signal.

                    ' Save record to file
                    ' If saveData = True Then

                    ' Find name and scale factor of this channel
                    Select Case channelId
                        Case CHANNEL_A
                            channelName = "A"
                            'UPGRADE_WARNING: Couldn't resolve default property of object scaleValueChA. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            scaleValue = scaleValueChA
                        Case CHANNEL_B
                            channelName = "B"
                            scaleValue = scaleValueChB
                        Case Else
                            MsgBox("Error: Invalid channelId " & channelId)
                            Exit Sub
                    End Select

                    ' Delimit the start of this record in the file
                    '                    Print #fileHandle, "--> CH" & channelName & " record " & record + 1 & " begin"
                    '                    Print #fileHandle, ""

                    ' Write record samples to file
                    For sampleInRecord = 0 To samplesPerRecord - 1

                        ' Get a sample value from the buffer
                        ' Note that the digitizer returns 16-bit unsigned sample values
                        ' that are stored in a 16-bit signed integer array, so convert
                        ' a signed 16-bit value to unsigned.
                        If (buffer(sampleInRecord) < 0) Then
                            sampleValue = buffer(sampleInRecord) + 65536
                        Else
                            sampleValue = buffer(sampleInRecord)
                        End If

                        ' Convert the sample value to volts
                        sampleVolts = scaleValue * (sampleValue - offsetValue)


                        ' Store new sampleVolts
                        ReDim Preserve arrWF(UBound(arrWF) + 1)
                        arrWF(sampleInRecord) = sampleVolts

                        ' Write the sample value in volts to file
                        ' Print #33, Str(sampleInRecord) & vbTab & sampleVolts
                        ' This is what I think displays units
                        ' sampleInRecord runs from 0 to 1023
                        ' Sampling rate is 5000000 Hz
                        ' To get seconds: sampleInRecord/5000000
                        ' Print #33, Str(sampleInRecord / 5000000) & " sec." & vbTab & Format(sampleVolts, "0.0000") & " V"
                        ' Print #33, Str(sampleInRecord) & vbTab & Format(sampleVolts, "0.0000")

                        PrintLine(44, Str(sampleInRecord / samplesPerSec) & " sec." & vbTab & VB6.Format(sampleVolts, "0.000000") & " V")

                    Next sampleInRecord

                    ' Ben: dummy placeholders
                    xinc = 0.1
                    trigpos = preTriggerSamples
                    vUnits = CStr(1)
                    hUnits = CStr(1)
                    xIncr = xinc
                    XTrig = trigpos
                    xMarker = XTrig
                    If IsArray(arrWF) Then ' check to be sure returned value is an array
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        zzzzzzz = UBound(arrWF) - LBound(arrWF) + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ToolTip1.SetToolTip(StatusL, "Scope Array Size = " & zzzzzzz)
                        'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        If zzzzzzz > 100000 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object zzzzzzz. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            MsgBox("Array too Large  size= " & zzzzzzz, , "Fix settings on scope")
                        End If
                        IDPoints = UBound(arrWF) - LBound(arrWF) + 1
                        'UPGRADE_WARNING: Couldn't resolve default property of object ccc. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        ccc = 0
                    Else
                        MsgBox("Error not array " & channel)
                    End If
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    jj = 0
                    ' Print #33, "Ch" & (channel + 1) & " " & xInc & " " & hUnits & " " & XTrig & " " & vUnits
                    For i = LBound(arrWF) To UBound(arrWF)
                        'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        'Histo(channel + 1, jj) = arrWF(i) * 1000000.0#
                        'Print #33, arrWF(i)
                        If record = 0 Then
                            'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            xDataRef(channel, jj) = arrWF(i)
                        Else
                            'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            xDataRef(channel, jj) = xDataRef(channel, jj) + arrWF(i)
                        End If
                        'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                        jj = jj + 1
                    Next i
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Max_Bins = jj

                    '                    ' Delimit the end of this record in the file
                    '                    Print #fileHandle, "<-- CH " & channelName & " record " & record + 1 & " end"
                    '                    Print #fileHandle, ""

                    ' End If  ' saveData

                    ' Draw record on screen
                    If DrawData = True Then
                        Call DrawRecord(buffer, samplesPerRecord, channelId)
                    End If

                    bytesTransferred = bytesTransferred + bytesPerRecord

                End If ' (channelMask And channelId) <> 0

            Next channel

            ' PictureBox.AutoRedraw = True
            ' Save picture
            ' SavePicture PictureBox.Image, "C:\Example.bmp"
            'UPGRADE_WARNING: SavePicture was upgraded to System.Drawing.Image.Save and has a new behavior. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="9B7D5ADD-D8FE-4819-A36C-6DEDAF088CC7"'
            PictureBox.Image.Save(ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".noise.bmp")
            ' Convert from bmp to jpeg
            ' BmpToJpeg "C:\Example" & ".bmp", "C:\Example" & ".jpg", 100
            BmpToJpeg(ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".noise.bmp", ImagePath & "Run" & VB6.Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".noise.jpg", 100)

            ' If the abort button was pressed, then stop reading records

            If GetInputState <> 0 Then System.Windows.Forms.Application.DoEvents()
            'UPGRADE_WARNING: Couldn't resolve default property of object abortAcquire. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'

        Next record





        ' Average out the noise samples
        'Dim ii As Integer
        'Dim kk As Integer
        'For ii = LBound(xDataRef, 1) To UBound(xDataRef, 1)
        '    For kk = LBound(xDataRef, 2) To UBound(xDataRef, 2)
        '       xDataRef(ii, kk) = xDataRef(ii, kk) / recordsPerCapture
        '       Next kk
        '    End
        '    Next ii
        'End

        'For channel = LBound(xDataRef, 1) To UBound(xDataRef, 1)
        For channel = 0 To channelsPerBoard - 1
            For i = 0 To preTriggerSamples + postTriggerSamples - 1
                'For i = LBound(xDataRef, 2) To UBound(xDataRef, 2)
                xDataRef(channel, i) = xDataRef(channel, i) / recordsPerCapture
                PrintLine(33, Str(i / samplesPerSec) & " sec." & vbTab & VB6.Format(xDataRef(channel, i), "0.000000") & " V")
            Next i
        Next channel


        ' Smooth out the noise samples
        For i = 0 To 1 'now create smoothed data
            'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xsumref = 0
            For ii = 0 To ISmooth - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xsumref = xsumref + xDataRef(i, ii)
            Next ii
            For ii = 0 To ISmooth / 2 - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xDataRef(i + 4, ii) = xsumref / ISmooth
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
            Next ii
            'xData(i + 4) = xsum / ISmooth

            For ii = ISmooth / 2 To Max_Bins - ISmooth / 2 - 1
                'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xsumref = 0
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                For jj = ii - ISmooth / 2 To ii + ISmooth / 2
                    'UPGRADE_WARNING: Couldn't resolve default property of object jj. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    xsumref = xsumref + xDataRef(i, jj)
                Next jj
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xDataRef(i + 4, ii) = xsumref / (ISmooth + 1)
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
                'xsum = xsum - xData(i, ii - ISmooth) + xData(i, ii)
                'xData(i + 4, ii - ISmooth) = xsum / ISmooth
                'Histo(i + 5, ii - ISmooth) = 1000000# * xsum / ISmooth
            Next ii
            For ii = Max_Bins - ISmooth / 2 To Max_Bins
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xsumref. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xDataRef(i + 4, ii) = xsumref / ISmooth
                'UPGRADE_WARNING: Couldn't resolve default property of object ii. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'Histo(i + 5, ii) = 1000000.0# * xData(i + 4, ii)
            Next ii
        Next i


        Call DrawAveragedSmoothedRecord(xDataRef, preTriggerSamples + postTriggerSamples, channelsPerBoard)



        FileClose(33)

        FileClose(44)

        '    If Check1(0).value = 1 Then Tvc1.WriteString "SAVe:WAVEform CH1, REF1"
        '    If Check1(1).value = 1 Then Tvc1.WriteString "SAVe:WAVEform CH2, REF2"
        'save image
        '    Dim strID As String
        '    strID = "EXPORT:FILEName """ & ImagePath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".noise.jpg"""
        '    Tvc1.WriteString strID
        '    strID = "EXPort:FORMat JPEG"
        '    Tvc1.WriteString strID
        '    strID = "EXPort:IMAGE NORMal"
        '    Tvc1.WriteString strID
        '    strID = "EXPort STARt"
        '    Tvc1.WriteString strID
        ' 32 = 2^5, maintains veto for TPC for bit D5
        ' Out Val("&H378"), IPrM + HVOn + Data7On + 32 'turn on pulser and HV
        Dim IsInline As Short
        IsInline = 0
        If IPrM = 0 Then IsInline = 32
        Out(Val("&H378"), IPrM + IsInline + HVOn + Data7On) 'turn on pulser and HV, no inhibit this time since there is no TPC

    End Sub

    Private Sub Timer1_Tick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Timer1.Tick
        'Static idir As Short
        'If idir = 0 Then
        '    Label2(0).Left = Label2(0).Left + 1
        '    If Label2(0).Left > 470 Then idir = 1
        'Else
        '    Label2(0).Left = Label2(0).Left - 1
        '    If Label2(0).Left < 47 Then idir = 0
        'End If
        ' The values commented out are in twips
        'Const E_RADIUS As Single = 20
        'Const O_Radius As Single = 1240
        ' The electron and electron orbot radii in pixels
        Const E_RADIUS As Single = 2
        Const O_Radius As Single = 83

        Dim cx As Single
        Dim cy As Single
        Dim X As Single
        Dim Y As Single

        Dim blackPen As New Pen(Color.Black)



        ' Draw the orbits and electrons. cx and cy are the center of the atom
        ' These commented out values are in twips
        'cx = 5240
        'cy = 840
        ' These values are in pixels
        cx = 351
        cy = 58

        EllipsePoint(X, Y, m_Alpha * 10.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 0 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(blackPen, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)


        EllipsePoint(X, Y, -m_Alpha * 11.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 12.0# / 30.0#, O_Radius, 0.5, 0.15, 2 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 13.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 3 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 14.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 4 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 15.0# / 30.0#, O_Radius, 0.5, 0.15, 5 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 16.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 6 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 17.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 7 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 18.0# / 30.0#, O_Radius, 0.5, 0.15, 8 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 19.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 9 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 20.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 10 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 21.0# / 30.0#, O_Radius, 0.5, 0.15, 11 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 22.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 12 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 23.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 13 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 24.0# / 30.0#, O_Radius, 0.5, 0.15, 14 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 25.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 15 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, m_Alpha * 26.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 16 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        EllipsePoint(X, Y, -m_Alpha * 27.0# / 30.0#, O_Radius, 0.5, 0.15, 17 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = &H8000000F
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, &H8000000F
        g.DrawEllipse(Pens.Black, X - E_RADIUS, Y - E_RADIUS, 2 * E_RADIUS, 2 * E_RADIUS)

        ' Draw the nucleus.
        '    Me.FillColor = vbBlack
        '    Me.Circle (cx, cy), 80, vbBlack

        ' Rotate a bit.
        m_Alpha = m_Alpha + DALPHA
        ' Draw the orbits and electrons.
        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 0 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, m_Alpha * 10.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 0 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 1 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, -m_Alpha * 11.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 1 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        EllipsePoint(X, Y, m_Alpha * 12.0# / 30.0#, O_Radius, 0.5, 0.15, 2 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 3 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, -m_Alpha * 13.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 3 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 4 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, m_Alpha * 14.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 4 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius, 0.5, 0.15, 5 * PI / 18, cx, cy, System.Drawing.Color.Blue)
        EllipsePoint(X, Y, -m_Alpha * 15.0# / 30.0#, O_Radius, 0.5, 0.15, 5 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 6 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, m_Alpha * 16.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 6 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))
        ' new for Argon


        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 7 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, -m_Alpha * 17.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 7 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius, 0.5, 0.15, 8 * PI / 18, cx, cy, System.Drawing.Color.Blue)
        EllipsePoint(X, Y, m_Alpha * 18.0# / 30.0#, O_Radius, 0.5, 0.15, 8 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 9 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, -m_Alpha * 19.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 9 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 10 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, m_Alpha * 20.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 10 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius, 0.5, 0.15, 11 * PI / 18, cx, cy, System.Drawing.Color.Blue)
        EllipsePoint(X, Y, -m_Alpha * 21.0# / 30.0#, O_Radius, 0.5, 0.15, 11 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 12 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, m_Alpha * 22 / 30.0#, O_Radius * 1.2, 0.5, 0.15, 12 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 13 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, -m_Alpha * 23.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 13 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius, 0.5, 0.15, 14 * PI / 18, cx, cy, System.Drawing.Color.Blue)
        EllipsePoint(X, Y, m_Alpha * 24.0# / 30.0#, O_Radius, 0.5, 0.15, 14 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.2, 0.5, 0.15, 15 * PI / 18, cx, cy, System.Drawing.Color.Red)
        EllipsePoint(X, Y, -m_Alpha * 25.0# / 30.0#, O_Radius * 1.2, 0.5, 0.15, 15 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbRed
        g.FillEllipse(Brushes.Red, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius * 1.3, 0.5, 0.15, 16 * PI / 18, cx, cy, System.Drawing.Color.Lime)
        EllipsePoint(X, Y, m_Alpha * 26.0# / 30.0#, O_Radius * 1.3, 0.5, 0.15, 16 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Lime)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbGreen
        g.FillEllipse(Brushes.Lime, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))

        DrawOrbitEllipse(O_Radius, 0.5, 0.15, 17 * PI / 18, cx, cy, System.Drawing.Color.Blue)
        EllipsePoint(X, Y, -m_Alpha * 27.0# / 30.0#, O_Radius, 0.5, 0.15, 17 * PI / 18, cx, cy)
        'UPGRADE_ISSUE: Form property PrMF.FillColor was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.FillColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue)
        'UPGRADE_ISSUE: Form method PrMF.Circle was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'Me.Circle (X, Y), E_RADIUS, 0 'vbBlue
        g.FillEllipse(Brushes.Blue, New Rectangle(X - E_RADIUS, Y - E_RADIUS, X + E_RADIUS, Y + E_RADIUS))
        Static ffirst, green, red, blue, iixt As Object

        '
        'UPGRADE_WARNING: Couldn't resolve default property of object ffirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If ffirst <> 57 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object ffirst. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            ffirst = 57
            'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            red = 128
            'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            green = 128
            'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blue = 128
        End If

        'UPGRADE_WARNING: Couldn't resolve default property of object iixt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iixt = iixt + 1
        'UPGRADE_WARNING: Couldn't resolve default property of object iixt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If iixt <> 3 Then Exit Sub
        'Beep
        'UPGRADE_WARNING: Couldn't resolve default property of object iixt. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        iixt = 0
        If Rnd() > 0.5 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            red = red + 10
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            red = red - 10
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If red < 20 Then red = 20
        'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If red > 225 Then red = 225

        If Rnd() > 0.5 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blue = blue + 10
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            blue = blue - 10
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If blue < 20 Then blue = 20
        'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If blue > 225 Then blue = 225

        If Rnd() > 0.5 Then
            'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            green = green + 10
        Else
            'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            green = green - 10
        End If
        'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If green < 20 Then green = 20
        'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If green > 225 Then green = 225

        'UPGRADE_WARNING: Couldn't resolve default property of object blue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object green. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object red. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        Label2(2).ForeColor = System.Drawing.ColorTranslator.FromOle(RGB(red, green, blue))

    End Sub

    'UPGRADE_WARNING: Event TimeT.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub TimeT_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles TimeT.TextChanged
        KeyHit = 1
        KeyTimer = 0
        Exit Sub
        IRate = Val(TimeT.Text)
        If IRate < GetNPrM() Then IRate = GetNPrM()
        If xRate > IRate * 60 Then xRate = IRate * 60
    End Sub

    Private Function GetNPrM() As Short
        Dim i, j As Short
        i = 0
        For j = 0 To NPrM - 1
            If Check4(j).CheckState = 1 And Check4(j).Enabled = True Then i = i + 1
        Next j
        GetNPrM = i
    End Function

    Private Sub TakeData()

        RunNumL.Text = VB6.Format(iiRun, "#0000") & "_" & VB6.Format(IPrM, "00")
        DataFileName = DataFilePath & "Run_" & VB6.Format(iiRun, "000000") & "_" & VB6.Format(IPrM, "00") & ".txt"
        AllTracesFileName = AllTracesPath & "Run_" & VB6.Format(iiRun, "000000") & "_" & VB6.Format(IPrM, "00") & ".txt"
        RunFileL.Text = DataFileName

        PulserWait.Interval = 60000
        PulserWait.Enabled = True

        List3.Items.Add("Configuring board")

        ' Provided device checks out, configure the Alazar board, Ben
        ' Check first to see if the current device is really connected to something
        Dim systemID As Object
        Dim boardID As Integer
        Dim boardHandle As Integer
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        systemID = 1
        boardID = 1
        'UPGRADE_WARNING: Couldn't resolve default property of object systemID. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        boardHandle = AlazarGetBoardBySystemID(systemID, boardID)
        'UPGRADE_WARNING: Couldn't resolve default property of object zlocal. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        If (boardHandle = 0) Then
            ' Oops, the device isn't available, notify the user and exit
            MsgBox("BoardID " & boardID & " not available." & Chr(10) & "You need to select an available device" & Chr(10) & "for the Alazar boardID in the VB IDE.", MsgBoxStyle.OkOnly, "Startup Error")
            End
        End If
        Dim result As Boolean
        result = ConfigureBoard(boardHandle)
        If (result <> True) Then
            Exit Sub
        End If

        List3.Items.Add("Board configured")

        '   set horizontal and vertical scales
        '    Dim strID As String
        '
        '    If IPrM = 0 Or IPrM = 2 Or IPrM = 3 Then ' long PrM
        '            strID = "HORizontal:MAIn:SCAle " & Text6.text
        '    Else ' short PrM
        '            strID = "HORizontal:MAIn:SCAle " & Text5.text
        '    End If
        '    Tvc1.WriteString strID

        'change V scale
        '    If IPrM = 0 Then
        '            strID = "CH2:SCAle 20E-03"
        '    Else
        '            strID = "CH2:SCAle 10E-03"
        '    End If
        ' Reconfigure vertical scale for each purity monitor, will likely be needed for Alazar, not sure yet
        'strID = "CH1:SCAle " & Text8(IPrM).text & "E-03"
        'Tvc1.WriteString strID
        'strID = "MATH1:VERTical:SCAle " & Text8(IPrM).text & "E-03"
        'Tvc1.WriteString strID
        'strID = "CH2:SCAle " & Text7(IPrM).text & "E-03"
        'Tvc1.WriteString strID
        'strID = "MATH2:VERTical:SCAle " & Text7(IPrM).text & "E-03"
        'Tvc1.WriteString strID

        ' HVOnStall.Enabled = True

        ' Tells parallel port to turn off HV and any running flash lamps
        ' 32 = 2^5, activates veto for TCP for bit D5
        ' Out Val("&H378"), 8 + Data7On + 32 'turn on pulser

        ' Sleep 10000

        ' Tells parallel port to turn on the flash lamp for each PM
        ' 32 = 2^5, activates veto for TCP for bit D5
        ' Out Val("&H378"), IPrM + Data7On + 32 'turn on pulser

        Dim IsInline As Short
        IsInline = 0
        If IPrM = 0 Then IsInline = 32
        Out(Val("&H378"), IPrM + IsInline + Data7On) 'turn on pulser, no longer need to worry about TPC veto

        StatusL.Text = "Pulser Turned ON, PrM = " & IPrM & " Pass = " & PassCnt
        '    PassCnt = 0

        HVWait.Interval = 30000
        HVWait.Enabled = True


    End Sub

    ' Configure timebase, trigger, and input parameters
    Private Function ConfigureBoard(ByRef boardHandle As Integer) As Boolean
        Dim retCode As Integer
        Dim samplesPerSec As Double
        Dim sampleRateId As Integer
        Dim triggerDelay_sec As Double
        Dim triggerDelay_samples As Integer
        Dim triggerTimeout_sec As Double
        Dim triggerTimeout_clocks As Integer

        ' Set default return value to indicate failure
        ConfigureBoard = False

        ' Select clock parameters as required
        ' 0, 2, 3 want 1 MSPS
        ' 1 and 4 want 5 MSPS
        ' Note on 4/14/2014, PrM 0 is short!
        ' Note on 4/15/2014, the 5 MSPS rate doesn't work for 0, I don't know why
        If IPrM = 0 Or IPrM = 2 Or IPrM = 3 Then
            'If IPrM = 2 Or IPrM = 3 Then
            sampleRateId = SAMPLE_RATE_2MSPS
            samplesPerSec = 2000000
        Else
            sampleRateId = SAMPLE_RATE_5MSPS
            samplesPerSec = 5000000
        End If

        retCode = AlazarSetCaptureClock(boardHandle, INTERNAL_CLOCK, sampleRateId, CLOCK_EDGE_RISING, 0)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetCaptureClock failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Select CHA input parameters as required
        ' InputRangeIdChA = INPUT_RANGE_PM_2_V, now set as combobox
        retCode = AlazarInputControl(boardHandle, CHANNEL_A, DC_COUPLING, InputRangeIdChA, IMPEDANCE_1M_OHM)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarInputControl CHA failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Select CHB input parameters as required
        ' InputRangeIdChB = INPUT_RANGE_PM_2_V, now set as combobox
        retCode = AlazarInputControl(boardHandle, CHANNEL_B, DC_COUPLING, InputRangeIdChB, IMPEDANCE_1M_OHM)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarInputControl CHB failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Select trigger inputs and levels as required
        retCode = AlazarSetTriggerOperation(boardHandle, TRIG_ENGINE_OP_J, TRIG_ENGINE_J, TRIG_EXTERNAL, TRIGGER_SLOPE_POSITIVE, 200, TRIG_ENGINE_K, TRIG_DISABLE, TRIGGER_SLOPE_POSITIVE, 128)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetTriggerOperation failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Select external trigger parameters as required
        retCode = AlazarSetExternalTrigger(boardHandle, DC_COUPLING, ETR_5V)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetExternalTrigger failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Set trigger delay as required
        triggerDelay_sec = 0
        triggerDelay_samples = triggerDelay_sec * samplesPerSec + 0.5
        retCode = AlazarSetTriggerDelay(boardHandle, triggerDelay_samples)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetTriggerDelay failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' TODO: Set trigger timeout as required

        ' NOTE:
        ' The board will wait for a for this amount of time for a trigger event.
        ' If a trigger event does not arrive, then the board will automatically
        ' trigger. Set the trigger timeout value to 0 to force the board to wait
        ' forever for a trigger event.

        ' IMPORTANT:
        ' The trigger timeout value should be set to zero after appropriate
        ' trigger parameters have been determined, otherwise the
        ' board may trigger if the timeout interval expires before a
        ' hardware trigger event arrives.

        triggerTimeout_sec = 0
        triggerTimeout_clocks = triggerTimeout_sec / 0.00001 + 0.5
        retCode = AlazarSetTriggerTimeOut(boardHandle, triggerTimeout_clocks)
        If (retCode <> ApiSuccess) Then
            MsgBox("Error: AlazarSetTriggerTimeOut failed -- " & AlazarErrorToText(retCode))
            Exit Function
        End If

        ' Attempt to setup board to take multiple records
        ' retCode = AlazarSetRecordCount(boardHandle, 10)
        ' If (retCode <> ApiSuccess) Then
        '    MsgBox ("Error: AlazarSetRecord failed -- " & AlazarErrorToText(retCode))
        '    Exit Function
        'End If

        ' Attempt to setup board to average records
        'retCode = AlazarSetParameter(boardHandle, CHANNEL_A, ACF_RECORDS_TO_AVERAGE, 10)
        'If (retCode <> ApiSuccess) Then
        '    MsgBox ("Error: AlazarSetParameter failed -- " & AlazarErrorToText(retCode))
        '    Exit Function
        'End If
        ' Attempt to setup board to average records
        'retCode = AlazarSetParameter(boardHandle, CHANNEL_B, ACF_RECORDS_TO_AVERAGE, 10)
        'If (retCode <> ApiSuccess) Then
        '    MsgBox ("Error: AlazarSetParameter failed -- " & AlazarErrorToText(retCode))
        '    Exit Function
        'End If


        ' Set return code to indicate success
        ConfigureBoard = True

    End Function


    ' Convert input range identifier to volts
    Private Function InputRangeIdToVolts(ByVal inputRangeId As Integer) As Double

        Select Case inputRangeId
            Case INPUT_RANGE_PM_40_MV
                InputRangeIdToVolts = 0.04
            Case INPUT_RANGE_PM_50_MV
                InputRangeIdToVolts = 0.05
            Case INPUT_RANGE_PM_80_MV
                InputRangeIdToVolts = 0.08
            Case INPUT_RANGE_PM_100_MV
                InputRangeIdToVolts = 0.1
            Case INPUT_RANGE_PM_200_MV
                InputRangeIdToVolts = 0.2
            Case INPUT_RANGE_PM_400_MV
                InputRangeIdToVolts = 0.4
            Case INPUT_RANGE_PM_500_MV
                InputRangeIdToVolts = 0.5
            Case INPUT_RANGE_PM_800_MV
                InputRangeIdToVolts = 0.8
            Case INPUT_RANGE_PM_1_V
                InputRangeIdToVolts = 1
            Case INPUT_RANGE_PM_2_V
                InputRangeIdToVolts = 2
            Case INPUT_RANGE_PM_4_V
                InputRangeIdToVolts = 4
            Case INPUT_RANGE_PM_5_V
                InputRangeIdToVolts = 5
            Case INPUT_RANGE_PM_8_V
                InputRangeIdToVolts = 8
            Case INPUT_RANGE_PM_10_V
                InputRangeIdToVolts = 10
            Case INPUT_RANGE_PM_20_V
                InputRangeIdToVolts = 20
            Case Else
                InputRangeIdToVolts = 0
        End Select

    End Function


    ' Draw record in picture box
    Private Sub DrawRecord(ByRef buffer() As Short, ByVal samplesPerRecord As Integer, ByVal channelId As Integer)
        Dim sampleInRecord As Integer
        Dim xLeft, xRight As Object
        Dim xRange As Integer
        Dim yBottom, yTop, yCenter As Object
        Dim yRange As Integer
        Dim sampleZero As Object
        Dim sampleRange As Integer
        Dim xOld As Object
        Dim yOld As Integer
        Dim color As System.Drawing.Color

        Dim bit As Bitmap = New Bitmap(PictureBox.Width, PictureBox.Height)
        Dim pictureGraphics As Graphics = Graphics.FromImage(bit)


        ' Set horizontal range
        'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xLeft = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object xRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xRight = PictureBox.ClientRectangle.Width
        'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xRange = xRight - xLeft

        ' Set vertical range
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yTop = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yBottom = PictureBox.ClientRectangle.Height
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yCenter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yCenter = (yBottom - yTop) / 2
        ' Was this
        ' yRange = (yBottom - yTop) / 10
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yRange = (yBottom - yTop) / 2

        ' Sample codes are stored in the most significant bits of each 16-bit sample value
        'UPGRADE_WARNING: Couldn't resolve default property of object sampleZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        sampleZero = 32768
        sampleRange = 32767

        ' Set record color
        If channelId = CHANNEL_A Then
            color = System.Drawing.Color.Blue
        Else
            color = System.Drawing.Color.Red
        End If

        Dim tracePen As New System.Drawing.Pen(color)


        Dim sampleValue As Integer
        Dim X As Object
        Dim Y As Integer
        For sampleInRecord = 0 To samplesPerRecord - 1

            ' The digitizer returns an array of 16-bit unsigned sample values
            ' that are stored in VB 16-bit signed integer array, so convert
            ' signed 16-bit data to unsigned
            If (buffer(sampleInRecord) < 0) Then
                sampleValue = buffer(sampleInRecord) + 65536
            Else
                sampleValue = buffer(sampleInRecord)
            End If

            ' Find the coordinates of this sample value
            'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            X = xLeft + xRange * sampleInRecord / samplesPerRecord
            'UPGRADE_WARNING: Couldn't resolve default property of object sampleZero. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object yCenter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            Y = yCenter - yRange * 0.5 * (sampleValue - sampleZero) / sampleRange

            ' Draw a line segment from the previous sample coordinates to the
            ' current sample coordinates
            If sampleInRecord > 0 Then
                'UPGRADE_ISSUE: PictureBox method PictureBox.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                'PictureBox.Line (xOld, yOld) - (X, Y), color
                g.DrawLine(tracePen, xOld, yOld, X, Y)
            End If

            ' Save this coordinate as the start point of the next line segment
            'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Couldn't resolve default property of object xOld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            xOld = X
            yOld = Y

        Next sampleInRecord

        ' PictureBox.AutoRedraw = True
        ' Save picture
        ' SavePicture PictureBox.Image, "C:\Example.bmp"
        ' Convert from bmp to jpeg
        ' BmpToJpeg "C:\Example" & ".bmp", "C:\Example" & ".jpg", 100

    End Sub




    'Draw smoothed, averaged traces
    Private Sub DrawAveragedSmoothedRecord(ByRef xDataToDraw(,) As Single, ByVal samplesPerRecord As Integer, ByVal channelsPerBoard As Integer)
        Dim i As Object
        Dim channel As Object

        Dim sampleInRecord As Integer
        Dim xLeft, xRight As Object
        Dim xRange As Integer
        Dim yBottom, yTop, yCenter As Object
        Dim yRange As Integer
        Dim sampleZero As Integer
        Dim sampleRange As Single
        Dim xOld As Object
        Dim yOld As Single
        'Dim color As Integer

        Dim color As System.Drawing.Color
        Dim bit As Bitmap = New Bitmap(PictureBox.Width, PictureBox.Height)
        Dim pictureGraphics As Graphics = Graphics.FromImage(bit)

        ' Set horizontal range
        'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xLeft = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object xRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xRight = Picture1.ClientRectangle.Width
        'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object xRight. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        xRange = xRight - xLeft

        ' Set vertical range
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yTop = 0
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yBottom = Picture1.ClientRectangle.Height
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yCenter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yCenter = (yBottom - yTop) / 2
        ' Was this
        ' yRange = (yBottom - yTop) / 10
        'UPGRADE_WARNING: Couldn't resolve default property of object yTop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'UPGRADE_WARNING: Couldn't resolve default property of object yBottom. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        yRange = (yBottom - yTop) / 2




        Dim sampleValue As Single
        Dim X As Object
        Dim Y As Single
        For channel = 0 To channelsPerBoard - 1

            ' Set record color
            'UPGRADE_WARNING: Couldn't resolve default property of object channel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If channel + 1 = CHANNEL_A Then
                sampleRange = InputRangeIdToVolts(InputRangeIdChA)
                color = System.Drawing.Color.Blue
            Else
                sampleRange = InputRangeIdToVolts(InputRangeIdChB)
                color = System.Drawing.Color.Red
            End If

            Dim tracePen As New System.Drawing.Pen(color)

            For i = 0 To samplesPerRecord - 1

                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object channel. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                sampleValue = xDataToDraw(channel + 4, i)

                ' Find the coordinates of this sample value
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xLeft. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                X = xLeft + xRange * i / samplesPerRecord
                'UPGRADE_WARNING: Couldn't resolve default property of object yCenter. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                Y = yCenter - yRange * 0.5 * (sampleValue) / sampleRange
                'Y = yCenter

                ' Draw a line segment from the previous sample coordinates to the
                ' current sample coordinates
                'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If i > 0 Then
                    'UPGRADE_ISSUE: PictureBox method Picture1.Line was not upgraded. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
                    'Picture1.Line (xOld, yOld) - (X, Y), color
                    g.DrawLine(tracePen, xOld, yOld, X, Y)
                End If

                ' Save this coordinate as the start point of the next line segment
                'UPGRADE_WARNING: Couldn't resolve default property of object X. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Couldn't resolve default property of object xOld. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                xOld = X
                yOld = Y


            Next i
        Next channel



    End Sub



End Class
