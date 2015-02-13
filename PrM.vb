Option Strict Off
Option Explicit On
Module PrM
	Public Declare Function Inp Lib "inpout32.dll"  Alias "Inp32"(ByVal PortAddress As Short) As Short
	Public Declare Sub Out Lib "inpout32.dll"  Alias "Out32"(ByVal PortAddress As Short, ByVal value As Short)
	
	Public iiRun, IRate, iiFile As Integer
	Public xPath As String
	Public DataFilePath, DataFileName As String
	Public xRate As Integer
	Public IAtom As Short
	Public xIncr As Double
	Public XTrig As Integer
	Public PassCnt, LiquidWait, iRunning As Short
	Public xData(7, 100000) As Single '0-3 are raw, 4-7 are smoothed
	Public xDataRef(7, 100000) As Single
	Public IDPoints, ISmooth, IOk As Short
	Public Ichan2, Ichan1, Ichan3 As Short
	Public RMSCut As Single
	Public OneTrueLiquid, ISets As Short
	Public KeyHit, KeyTimer As Short
	Public LogPath, DataFileNameHist As String
	Public AllTracesPath, AllTracesFileName As String
	Public ImagePath As String
	Public isubtrch1 As Short
	Public IPrM, NPrM As Short
	
	
	Sub Anal_Data()
		Dim zztime As Object
        Dim LifeTime As Double = 0
		Dim AnoF As Object
		Dim Cathf As Object
		Dim RC As Object
        Dim va2 As Double = 0
        Dim ta2 As Double = 0
        Dim va1 As Double = 0
        Dim ta1 As Double = 0
		Dim da2 As Object
		Dim da1 As Object
		Dim a2 As Object
		Dim a1 As Object
		Dim xrise As Object
		Dim istop As Object
		Dim iix As Object
		Dim xAnoTime As Object
		Dim iAnotime As Object
		Dim xrms As Object
		Dim xsq As Object
		Dim xsum As Object
		Dim xCatTime As Object
        Dim ixmax As Integer = -1
		Dim i As Object
        Dim XBASE As Double = 0
		Dim IXTIME As Object
		Dim XMAX As Object
		Dim Istart As Integer
		Dim AnoPeak, CatPeak, AnoTrue, CatTrue, CatBase, AnoBase As Object
		Dim CatTime, DiodeTime, DiodePeak, DiodeBase, AnoRise As Object
		Dim AnoTime As Object
		Dim iCatTime As Object
		Dim Vch1 As Single
		On Error GoTo 0
		On Error Resume Next
		FileOpen(7, DataFilePath & "AnalyzedData.txt", OpenMode.Append)
		FileOpen(77, DataFilePath & "CondensedData.txt", OpenMode.Append)
		PrMF.List1.Items.Clear()
		PrMF.List2.Items.Clear()
		PrintLine(7, Today & " " & TimeOfDay & "," & iiRun & "," & IPrM & "," & PassCnt)
		Print(77, Today & " " & TimeOfDay & "," & iiRun & ",", TAB)
		PrMF.List1.Items.Add(Today & " " & TimeOfDay)
		PrMF.List1.Items.Add(" Run = " & iiRun & " Pass = " & PassCnt & " PrM = " & IPrM)
		' do diode (trigger) information
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XMAX = 1000
		'For i = 0 To IDPoints - 1
		'    If xData(Ichan3, i) < XMAX Then
		'        XMAX = xData(Ichan3, i)
		'        ixmax = i
		'    End If
		'Next i
		'DiodePeak = XMAX
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IXTIME = 0
		'IXTIME = -XTrig * xIncr + ixmax * xIncr
		'xsum = 0
		'For i = ixmax / 2 - 25 To ixmax / 2 + 24
		'    xsum = xsum + xData(Ichan3, i)
		'Next i
		'XBASE = xsum / 50#
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XBASE = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "Diode Peak,Time,Baseline," & Format(XMAX, "0.000e-00") & "," & Format(IXTIME, "0.000e-00") & "," & Format(XBASE, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List1.Items.Add("Diode Time = " & Format(IXTIME, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DiodeTime = IXTIME
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodePeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DiodePeak = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		DiodeBase = 0
		
		'   first get the anode time for noise removal, this will be repeated later
		'   wait 100 microseconds after the trigger to look for the anode max, T.Y. 12/01/11
		' Istart = XTrig + 0.0001 / xIncr
		'        Istart = XTrig
		
		'    XMAX = -1000
		'    For i = Istart To IDPoints - 1
		'        If xData(Ichan1, i) > XMAX Then
		'            XMAX = xData(Ichan1, i)
		'            ixmax = i
		'        End If
		'    Next i
		'    IXTIME = -XTrig * xIncr + ixmax * xIncr
		'    iAnoTime = ixmax
		'    'MsgBox (iAnoTime & " " & IXTIME & " " & XMAX)
		'    'MsgBox ("Check6.Value = " & Check6.Value)
		'    'MsgBox ("hello!")
		'    'MsgBox (isubtrch1)
		' now do cathode info
		'   wait 25 microseconds after the trigger to look for the cathode min and anode max
		'    Istart = XTrig + 0.000025 / xIncr
		Istart = XTrig + 0.00001 / xIncr
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XMAX = 1000
		For i = Istart To IDPoints - 1
			'For i = Istart To iAnoTime - 0.0001 / xIncr
			'        Vch1 = 0
			'        If isubtrch1 = 1 Then   ' subtract ch1 to remove noise
			'            Vch1 = xData(Ichan1, i)
			'        End If
			'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If xData(Ichan2, i) - xDataRef(Ichan2, i) < XMAX Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				XMAX = xData(Ichan2, i) - xDataRef(Ichan2, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ixmax = i
			End If
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object iCatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iCatTime = ixmax
		'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IXTIME = -XTrig * xIncr + ixmax * xIncr
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xCatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xCatTime = IXTIME
		'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xsum = 0
		For i = XTrig / 2 - 25 To XTrig / 2 + 24
			Vch1 = 0
			'        If isubtrch1 = 1 Then   ' subtract ch1 to remove noise
			'            Vch1 = xData(Ichan1, i)
			'        End If
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xsum = xsum + xData(Ichan2, i) - xDataRef(Ichan2, i) ' - Vch1
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XBASE = xsum / 50#
		'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xsq = 0
		For i = XTrig / 2 - 25 To XTrig / 2 + 24
			'        Vch1 = 0
			'        If isubtrch1 = 1 Then   ' subtract ch1 to remove noise
			'            Vch1 = xData(Ichan1, i)
			'        End If
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xsq = xsq + (XBASE - xData(Ichan2, i) + xDataRef(Ichan2, i)) ^ 2
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xrms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xrms = (xsq / 50#) ^ 0.5
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CatPeak = XMAX
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CatBase = XBASE
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "Cathode Peak,Time,Baseline," & Format(XMAX, "0.000e-00") & "," & Format(IXTIME, "0.000e-00") & "," & Format(XBASE, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List1.Items.Add("Cathode Peak = " & Format(XMAX, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List1.Items.Add("Cathode Time = " & Format(IXTIME, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List1.Items.Add("Cathode Baseline = " & Format(XBASE, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CatTime = IXTIME
		'UPGRADE_WARNING: Couldn't resolve default property of object xrms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If System.Math.Abs(CatPeak - CatBase) < RMSCut * xrms Then CatPeak = CatBase
		
		
		'now do anode
		'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xsum = 0
		For i = XTrig / 2 - 25 To XTrig / 2 + 24
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xsum = xsum + xData(Ichan1, i) - xDataRef(Ichan1, i)
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object xsum. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XBASE = xsum / 50#
		'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xsq = 0
		For i = XTrig / 2 - 25 To XTrig / 2 + 24
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xsq = xsq + (XBASE - xData(Ichan1, i) + xDataRef(Ichan1, i)) ^ 2
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object xsq. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xrms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xrms = (xsq / 50#) ^ 0.5
		'   wait 100 microseconds after the trigger to look for the anode max, T.Y. 12/01/11
		Istart = XTrig + 0.00001 / xIncr
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XMAX = -1000
		For i = Istart To IDPoints - 1
			'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If xData(Ichan1, i) - xDataRef(Ichan1, i) > XMAX Then
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				XMAX = xData(Ichan1, i) - xDataRef(Ichan1, i)
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				ixmax = i
			End If
		Next i
		'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object iAnotime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iAnotime = ixmax
		'UPGRADE_WARNING: Couldn't resolve default property of object ixmax. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		IXTIME = -XTrig * xIncr + ixmax * xIncr
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xAnoTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xAnoTime = IXTIME
		'iix = 0.1 * (iAnotime - iCatTime + 0.000005 / xIncr)
		'UPGRADE_WARNING: Couldn't resolve default property of object iCatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object iAnotime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object iix. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		iix = iAnotime - 0.3333 * (iAnotime - iCatTime + 0.000005 / xIncr)
		' MsgBox iAnotime & " " & iCatTime & " " & iix
		'UPGRADE_WARNING: Couldn't resolve default property of object iix. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		XBASE = xData(Ichan1, iix) - xDataRef(Ichan1, iix)
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoPeak = XMAX
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoBase = XBASE
		'UPGRADE_WARNING: Couldn't resolve default property of object istop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		istop = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		xrise = 0
		'UPGRADE_WARNING: Couldn't resolve default property of object xrms. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If System.Math.Abs(AnoPeak - AnoBase) < RMSCut * xrms Then
			'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			AnoPeak = AnoBase
			'UPGRADE_WARNING: Couldn't resolve default property of object istop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			istop = 1
		End If
		'UPGRADE_WARNING: Couldn't resolve default property of object istop. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If istop = 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			a1 = XBASE + 0.25 * (XMAX - XBASE)
			'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			a2 = XBASE + 0.75 * (XMAX - XBASE)
			'UPGRADE_WARNING: Couldn't resolve default property of object da1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			da1 = 10000#
			'UPGRADE_WARNING: Couldn't resolve default property of object da2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			da2 = 10000#
			'UPGRADE_WARNING: Couldn't resolve default property of object iAnotime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object iix. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			For i = iix To iAnotime
				'UPGRADE_WARNING: Couldn't resolve default property of object da1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(xData(Ichan1, i) - xDataRef(Ichan1, i) - a1) < da1 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object a1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object da1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					da1 = System.Math.Abs(xData(Ichan1, i) - xDataRef(Ichan1, i) - a1)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ta1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ta1 = -XTrig * xIncr + i * xIncr
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object va1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					va1 = xData(Ichan1, i) - xDataRef(Ichan1, i)
				End If
				'UPGRADE_WARNING: Couldn't resolve default property of object da2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				If System.Math.Abs(xData(Ichan1, i) - xDataRef(Ichan1, i) - a2) < da2 Then
					'UPGRADE_WARNING: Couldn't resolve default property of object a2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object da2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					da2 = System.Math.Abs(xData(Ichan1, i) - xDataRef(Ichan1, i) - a2)
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object ta2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					ta2 = -XTrig * xIncr + i * xIncr
					'UPGRADE_WARNING: Couldn't resolve default property of object i. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					'UPGRADE_WARNING: Couldn't resolve default property of object va2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
					va2 = xData(Ichan1, i) - xDataRef(Ichan1, i)
				End If
			Next i
			'UPGRADE_WARNING: Couldn't resolve default property of object va1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object va2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ta1. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object ta2. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			xrise = (ta2 - ta1) * (AnoPeak - AnoBase) / (va2 - va1)
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "Anode Peak,Time,Baseline,Rise," & Format(XMAX, "0.000e-00") & "," & Format(IXTIME, "0.000e-00") & "," & Format(XBASE, "0.000e-00") & "," & Format(xrise, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode Peak = " & Format(XMAX, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode Time = " & Format(IXTIME, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode Baseline = " & Format(XBASE, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode Rise = " & Format(xrise, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoRise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoRise = xrise
		'UPGRADE_WARNING: Couldn't resolve default property of object XMAX. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoPeak = XMAX
		'UPGRADE_WARNING: Couldn't resolve default property of object IXTIME. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoTime = IXTIME
		'UPGRADE_WARNING: Couldn't resolve default property of object XBASE. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoBase = XBASE
		
		'UPGRADE_WARNING: Couldn't resolve default property of object RC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		RC = 0.000119
		
		'UPGRADE_WARNING: Couldn't resolve default property of object RC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xCatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		Cathf = (xCatTime + 0.000006) / (RC * (1# - System.Math.Exp(-(xCatTime + 0.000006) / RC)))
		'UPGRADE_WARNING: Couldn't resolve default property of object RC. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object xrise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoF = (xrise + 0.000005) / (RC * (1# - System.Math.Exp(-(xrise + 0.000005) / RC)))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		AnoTrue = (AnoPeak - AnoBase) * AnoF
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		CatTrue = System.Math.Abs((CatPeak - CatBase) * Cathf)
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If AnoTrue > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			If (CatTrue / AnoTrue) > 0 Then
				'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object xAnoTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LifeTime = xAnoTime / (System.Math.Log(CatTrue / AnoTrue))
			Else
				'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
				LifeTime = 0
			End If
		End If
		
		'11/29/11 Set lifetime to 30 ms if it is greater than 30 ms or is negative. T.Yang
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LifeTime > 0.03 Or LifeTime < 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			LifeTime = 0.03
		End If
		
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "Cath Factor,Anode Factor,Anode True, Cathode True,LifeTime," & Format(Cathf, "0.000e-00") & "," & Format(AnoF, "0.000e-00") & "," & Format(AnoTrue, "0.000e-00") & "," & Format(CatTrue, "0.000e-00") & "," & Format(LifeTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodePeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(77, "Cath Factor,Anode Factor,Anode True, Cathode True,LifeTime,Diode Peak," & Format(Cathf, "0.000e-00") & "," & Format(AnoF, "0.000e-00") & "," & Format(AnoTrue, "0.000e-00") & "," & Format(CatTrue, "0.000e-00") & "," & Format(LifeTime, "0.000e-00") & "," & Format(DiodePeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Cath Factor = " & Format(Cathf, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode Factor = " & Format(AnoF, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Anode True = " & Format(AnoTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("Cathode True = " & Format(CatTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrMF.List2.Items.Add("LifeTime = " & Format(LifeTime, "0.000e-00"))
		On Error GoTo 0
		FileClose(7)
		FileClose(77)
		'Dim Istart As Long, AnoTrue, CatTrue, CatPeak, CatBase, AnoPeak, AnoBase
		'Dim DiodePeak, DiodeTime, DiodeBase, CatTime, AnoRise
		On Error GoTo LocalWrite
		'    Open LogPath & "Run" & iiRun & "." & iiFile & ".LogData.csv" For Output As #7
        FileOpen(7, LogPath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv", OpenMode.Output)
		On Error GoTo 0
		PrintLine(7, "[Data]")
		PrintLine(7, "Tagname,TimeStamp,Value")
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        zztime = Format(TimeOfDay, "hh:mm:ss")
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodePeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODEPEAK_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(DiodePeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODETIME_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(DiodeTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODEBASE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(DiodeBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHPEAK_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(CatPeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHTIME_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(CatTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHBASE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(CatBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEPEAK_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoPeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODETIME_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEBASE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoRise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODERISE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoRise, "0.000e-00"))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHFACTOR_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(Cathf, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEFACTOR_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoF, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODETRUE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(AnoTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHTRUE_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(CatTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_LIFETIME_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(LifeTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LifeTime > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(7, "LAPD.PRM_IMPURITIES_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(0.00015 / LifeTime, "0.000e-00"))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(7, "LAPD.PRM_IMPURITIES_" & IPrM & ".F_CV," & Today & " " & zztime & "," & Format(99999, "0.000e-00"))
		End If
		FileClose(7)
		
		'copy file to dropbox
		'    If Dir("C:\Dropbox\PrM\", vbDirectory) <> "" Then
		'        If Dir(LogPath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv") <> "" Then
		'            FileCopy LogPath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv", "C:\Dropbox\PrM\" & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv"
		'        Else
		'            Open DataFileNameHist For Append As #88
		'            Print #88, "-------" & Date & " " & Time & " " & LogPath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv does not exist"
		'            Close #88
		'        End If
		'        FileCopy DataFileNameHist, "C:\Dropbox\PrM\PrMLongHistory.txt"
		'    Else
		'        Open DataFileNameHist For Append As #88
		'        Print #88, "-------" & Date & " " & Time & " C:\Dropbox\PrM\ does not exist"
		'        Close #88
		'    End If
		
		iiFile = iiFile + 1
		
		Exit Sub
LocalWrite: 
		FileOpen(88, DataFileNameHist, OpenMode.Append)
		PrintLine(88, Today & " " & TimeOfDay & " Run Number = " & iiRun & " Remote Disk Unavailable")
		FileClose(88)
		
		'    Open DataFilePath & "Run" & iiRun & "." & iiFile & ".LogData.csv" For Output As #7
        FileOpen(7, DataFilePath & "Run" & Format(iiRun, "000000") & "." & IPrM & "." & iiFile & ".LogData.csv", OpenMode.Output)
		On Error GoTo 0
		iiFile = iiFile + 1
		PrintLine(7, "[Data]")
		PrintLine(7, "Tagname,TimeStamp,Value")
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        zztime = Format(TimeOfDay, "hh:mm:ss")
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodePeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODEPEAK.F_CV," & Today & " " & zztime & "," & Format(DiodePeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODETIME.F_CV," & Today & " " & zztime & "," & Format(DiodeTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object DiodeBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_DIODEBASE.F_CV," & Today & " " & zztime & "," & Format(DiodeBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHPEAK.F_CV," & Today & " " & zztime & "," & Format(CatPeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHTIME.F_CV," & Today & " " & zztime & "," & Format(CatTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHBASE.F_CV," & Today & " " & zztime & "," & Format(CatBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoPeak. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEPEAK.F_CV," & Today & " " & zztime & "," & Format(AnoPeak, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODETIME.F_CV," & Today & " " & zztime & "," & Format(AnoTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoBase. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEBASE.F_CV," & Today & " " & zztime & "," & Format(AnoBase, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoRise. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODERISE.F_CV," & Today & " " & zztime & "," & Format(AnoRise, "0.000e-00"))
		
		'UPGRADE_WARNING: Couldn't resolve default property of object Cathf. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHFACTOR.F_CV," & Today & " " & zztime & "," & Format(Cathf, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoF. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODEFACTOR.F_CV," & Today & " " & zztime & "," & Format(AnoF, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object AnoTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_ANODETRUE.F_CV," & Today & " " & zztime & "," & Format(AnoTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object CatTrue. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_CATHTRUE.F_CV," & Today & " " & zztime & "," & Format(CatTrue, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        PrintLine(7, "LAPD.PRM_LIFETIME.F_CV," & Today & " " & zztime & "," & Format(LifeTime, "0.000e-00"))
		'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
		If LifeTime > 0 Then
			'UPGRADE_WARNING: Couldn't resolve default property of object LifeTime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
			'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(7, "LAPD.PRM_IMPURITIES.F_CV," & Today & " " & zztime & "," & Format(0.00015 / LifeTime, "0.000e-00"))
		Else
			'UPGRADE_WARNING: Couldn't resolve default property of object zztime. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            PrintLine(7, "LAPD.PRM_IMPURITIES.F_CV," & Today & " " & zztime & "," & Format(99999, "0.000e-00"))
		End If
		
		FileClose(7)
	End Sub
End Module