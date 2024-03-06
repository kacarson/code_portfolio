Global g_Cycle As String
Option Base 1


Sub SSRSBillingSummaryQC3_2()
'
' SSRSBillingSummary Macro
'
Screenupdateing = False
Dim BillStrtDt As Integer
Dim findend As Integer
Dim FindDt As String
Dim Cycle As String
Dim NewDt As Date
Dim strt As Integer
Dim check22 As Double
Dim check23 As Double
Dim check24 As Double
Dim check25 As Double
Dim check26 As Double
Dim check27 As Double
Dim enddt As Integer
Dim O As Integer
Dim P As Integer
Dim CycleEnd As String
Dim CycleEndDt As Date
Dim CycleLength As Integer
Dim ProLength As Integer
Dim PropSS As String
Dim PropSSshort As String
Dim PropQC As String

PropQC = ActiveWorkbook.Name

'Finds cycle
BillStrtDt = InStr(Cells(1, 1), "C | ")
BillStrtDt = BillStrtDt + 4

findend = InStr(Cells(1, 1), " | D")

strt = BillStrtDt
BillStrtDt = findend - BillStrtDt
FindDt = Mid(Cells(1, 1), strt, BillStrtDt)
g_Cycle = FindDt
findend = InStr(FindDt, " - ")
FindDt = Mid(FindDt, 1, findend)
enddt = Len(g_Cycle)
O = findend + 3
P = enddt - 0
P = P - O + 1

CycleEnd = Right(g_Cycle, P)
CycleEndDt = CDate(CycleEnd)
NewDt = CDate(FindDt)
Cells(2, 1) = g_Cycle

'on error
Set UnMergeTarget = Columns("B:C")
UnMergeTarget.UnMerge
If Range("C2").Value = "" Then
    Columns("C").EntireColumn.Select
    selection.Delete Shift:=x1Left
End If
Set Target = Columns("B:C").Find(What:="Billable Subtotal", LookIn:=xlValues, LookAt:=xlWhole)
Target.EntireRow.Select
ActiveCell.Rows("1:1000").EntireRow.Select
selection.Delete Shift:=xlUp
Rows("1:1").Select
selection.Delete Shift:=xlUp
Rows("2:3").Select
selection.Delete Shift:=xlUp
Range("A1").Select
Cells.Select
Cells.EntireColumn.ColumnWidth = 9.43
Cells.EntireColumn.RowHeight = 13
Range("A1").Select
Rows("1:1").RowHeight = 38.25
Rows("1:1").Select
With selection
    .VerticalAlignment = xlBottom
    .HorizontalAlignment = xlLeft
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With
With ActiveWindow
    .SplitColumn = 0
    .SplitRow = 1
End With
ActiveWindow.FreezePanes = True
Cells.Select
Cells.EntireColumn.AutoFit
Range("A1").Select
ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
With selection
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlBottom
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
End With
selection.UnMerge
selection.EntireRow.Hidden = True
Range("A1").Select
Cells.Select
selection.UnMerge
Range("A1").Select

PropSS = findPropSpreadsheet(PropSSshort)

closeSS (PropSSshort)

'''SORT by MI Date
Range("E3").Select
ActiveWorkbook.Worksheets("Summary of Utilities Billed - Q").Sort.SortFields.Clear
ActiveWorkbook.Worksheets("Summary of Utilities Billed - Q").Sort.SortFields.Add Key:=Range("E3"), _
    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
    xlSortTextAsNumbers
With ActiveWorkbook.Worksheets("Summary of Utilities Billed - Q").Sort
    .SetRange Range("A3:DD800")
    .Header = xlNo
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

'insert MICK_Deposit
Dim ansWr As Integer
Dim answr2 As Integer
Dim answr3 As Integer
ansWr = MsgBox("Would you like to continue formatting?", vbYesNo)

'Yes for formatting block
If ansWr = 6 Then
    answr2 = MsgBox("Is this an Ontario Property?", vbYesNo)
    
'Per Unit Flat Fees
    Dim SDF As Integer
    Dim SDF2 As Integer
    Dim RegRR As Integer
    Dim ElecMeter As Integer 'Can serve as flat or metered fee
    Dim ElecMeter2 As Integer
    Dim ElecMeter3 As Integer
    Dim ElecMeter4 As Integer
    Dim ElecMeter5 As Integer
    Dim SSSadmin As Integer
    Dim WWA As Integer
    Dim WtrMtr As Integer
    Dim WtrMtr2 As Integer
    Dim WtrMtr3 As Integer
    Dim WtrMtr4 As Integer
    Dim WtrMtr5 As Integer
    Dim WtrMtr6 As Integer
    Dim WtrMtr7 As Integer
    Dim WtrMtr8 As Integer
    Dim WtrMtr9 As Integer
    Dim WtrMtr10 As Integer
    Dim WtrMtr11 As Integer
    Dim RegComp As Integer
    Dim ThrmMCol As Integer
    Dim ThermalMeter2_Col As Integer
    Dim ThermalMeter3_Col As Integer
    Dim ThermalMeter4_Col As Integer
    Dim ThermalMeter5_Col As Integer
    Dim ThermalMeter6_Col As Integer
    Dim ThermalMeter7_Col As Integer
    Dim ThermalMeter8_Col As Integer
    Dim ThrmAdCol As Integer
    Dim BadDtCol As Integer
    Dim COVIDRecCol As Integer
    Dim StormWaterSewerCol As Integer
    Dim GasAdmin_Col As Integer
    Dim GasMeter_Col As Integer
    Dim GasMeter2_Col As Integer
    Dim PaperBillTax_Col As Integer
    
    'RUBs Flat Fees
    Dim ElecCust As Integer
    Dim FireSup As Integer
    Dim Base As Integer
    Dim SewerBase As Integer
    Dim StormWaterDrainCol As Integer
    Dim GasCust_Col As Integer

    'Metered Fees
    ''Energy
    '''Charge
    Dim Deliv As Integer
    Dim Energy As Integer
    Dim Energy_2 As Integer
    Dim Energy_3 As Integer
    Dim LILO As Integer
    Dim Reg As Integer
    Dim MidPeak As Integer
    Dim OffPeak As Integer
    Dim OnPeak As Integer
    '''Consumption
    Dim Delivcons As Integer
    Dim EnCons As Integer
    Dim EnCons_2 As Integer
    Dim EnCons_3 As Integer
    Dim LILOcons As Integer
    Dim RegCons As Integer
    Dim MidPeakCons As Integer
    Dim OffPeakCons As Integer
    Dim OnPeakCons As Integer
    Dim ElecMeterCons As Integer
    '''Rate
    Dim RegChg_Rate() As Variant
    Dim Delivery_Rate() As Variant
    Dim Energy_Rate() As Variant
    Dim Energy2_Rate() As Variant
    Dim Energy3_Rate() As Variant
    Dim MidPeak_Rate() As Variant
    Dim OffPeak_Rate() As Variant
    Dim OnPeak_Rate() As Variant
    Dim ElecMeter_Rate() As Variant
    
    ''Water
    '''Charge
    Dim coldwtR As Integer
    Dim coldwtR2 As Integer
    Dim coldwtR3 As Integer
    Dim coldwtR4 As Integer
    Dim coldwtR5 As Integer
    Dim hotwtR As Integer
    Dim hotwtR2 As Integer
    Dim hotwtR3 As Integer
    Dim hotwtR4 As Integer
    Dim hotwtR5 As Integer
    Dim wtR As Integer
    Dim wtR2 As Integer
    Dim wtR3 As Integer
    Dim Sewer As Integer
    Dim Sewer2 As Integer
    '''Consumption
    Dim coldwtRcons As Integer
    Dim coldwtR2cons As Integer
    Dim coldwtR3cons As Integer
    Dim coldwtR4cons As Integer
    Dim coldwtR5cons As Integer
    Dim hotwtRcons As Integer
    Dim hotwtR2cons As Integer
    Dim hotwtR3cons As Integer
    Dim hotwtR4cons As Integer
    Dim hotwtR5cons As Integer
    Dim wtRcons As Integer
    Dim wtR2cons As Integer
    Dim wtR3cons As Integer
    Dim SewerCons As Integer
    Dim Sewer2Cons As Integer
    '''Rate
    Dim Water_Rate() As Variant
    Dim Water2_Rate() As Variant
    Dim Water3_Rate() As Variant
    Dim clDWater_Rate() As Variant
    Dim clDWater2_Rate() As Variant
    Dim clDWater3_Rate() As Variant
    Dim clDWater4_Rate() As Variant
    Dim clDWater5_Rate() As Variant
    Dim hTWater_Rate() As Variant
    Dim hTWater2_Rate() As Variant
    Dim hTWater3_Rate() As Variant
    Dim hTWater4_Rate() As Variant
    Dim hTWater5_Rate() As Variant
    Dim Sewer_Rate() As Variant
    Dim Sewer2_Rate() As Variant
    
    ''Thermal
    '''Charge
    Dim ThrmCol As Integer
    Dim Cooling_Col As Integer
    Dim Cooling2_Col As Integer
    Dim Cooling3_Col As Integer
    Dim Cooling4_Col As Integer
    Dim Heating_Col As Integer
    Dim Heating2_Col As Integer
    Dim Heating3_Col As Integer
    Dim Heating4_Col As Integer
    '''Consumption
    Dim ThrmCnCol As Integer
    Dim CoolingCon_Col As Integer
    Dim Cooling2Con_Col As Integer
    Dim Cooling3Con_Col As Integer
    Dim Cooling4Con_Col As Integer
    Dim HeatingCon_Col As Integer
    Dim Heating2Con_Col As Integer
    Dim Heating3Con_Col As Integer
    Dim Heating4Con_Col As Integer
    '''Rate
    Dim Thermal_Rate() As Variant
    Dim Cooling_Rate() As Variant
    Dim Heating_Rate() As Variant
    
    ''Gas
    Dim Gas_Col As Integer
    Dim GasCon_Col As Integer
    Dim Gas_Rate() As Variant
    
    'Taxes
    Dim SUPHST As Integer
    Dim HST As Integer
    Dim WyseHST As Integer
    Dim RegRbate As Integer
    Dim wyseRbate As Integer
    Dim GST As Integer
    Dim WyseGST As Integer
    Dim PST As Integer
    Dim WysePST As Integer
    Dim OESPHST As Integer
    Dim OESPRbate As Integer
    Dim OESPHSTCR As Integer
    Dim OESPRbateCR As Integer
    
    'One-Time and Miscellaneous
    Dim MI_DATE As Date
    Dim MIDATE As Integer
    Dim currow1 As Integer
    Dim QCCurRow As Integer
    Dim Dephigh As Currency
    Dim DepLow As Currency
    Dim UTILDEP As Integer
    Dim AutoPay As Integer
    Dim DepStat As Integer
    Dim DepAmt As Integer
    Dim AcctSUP As Integer
    Dim OESP As Integer
    Dim LAF As Double
    Dim x As Currency
    Dim XX As Integer
    Dim XY As Integer
    Dim HSTAccumulator
    Dim WyseHSTAccumulator
    Dim OERAccumulator
    Dim WyseOERAccumulator
    Dim GSTAccumulator
    Dim WyseGSTAccumulator
    Dim PSTAccumulator
    Dim WysePSTAccumulator
    Dim EbillAddie_Col As Integer
    Dim HasPyArr_Col As Integer
    Dim PAMonthly_Col As Integer
    Dim PrevBal_Col As Integer
    Dim PastDueArr_Col As Integer

      
    MI_DATE = NewDt
    Dephigh = Find_info("Deposit", 2, PropSS, PropSSshort)
    DepLow = Find_info("PAP_Deposit", 2, PropSS, PropSSshort)
    Setup_Fee = Find_info("Setup_Fee", 2, PropSS, PropSSshort)
    
    currow1 = 3
    MIDATE = 1
    UTILDEP = 1
    AutoPay = 1
    DepStat = 1
    DepAmt = 1
    AcctSUP = 1
    SUPHST = 1
    Delivcons = 1
    OESPHSTCR = 1
    OESPRbateCR = 1
    EnCons = 1
    EnCons_2 = 1
    EnCons_3 = 1
    LILOcons = 1
    RegCons = 1
    SDF = 9999
    SDF2 = 1
    RegRR = 1
    ElecMeter = 1
    ElecMeter2 = 1
    ElecMeter3 = 1
    ElecMeter4 = 1
    ElecMeter5 = 1
    SSSadmin = 1
    ElecCust = 1
    HST = 1
    WyseHST = 1
    RegRbate = 1
    wyseRbate = 1
    OESP = 1
    OESPHST = 1
    OESPRbate = 1
    GST = 1
    WyseGST = 1
    PST = 1
    WysePST = 1
    Deliv = 1
    Energy = 1
    Energy_2 = 1
    Energy_3 = 1
    LILO = 1
    Reg = 1
    QCCurRow = 3
    coldwtRcons = 1
    coldwtR2cons = 1
    coldwtR3cons = 1
    coldwtR4cons = 1
    coldwtR5cons = 1
    hotwtRcons = 1
    hotwtR2cons = 1
    hotwtR3cons = 1
    hotwtR4cons = 1
    hotwtR5cons = 1
    wtRcons = 1
    wtR2cons = 1
    wtR3cons = 1
    coldwtR = 1
    coldwtR2 = 1
    coldwtR3 = 1
    coldwtR4 = 1
    coldwtR5 = 1
    hotwtR = 1
    hotwtR2 = 1
    hotwtR3 = 1
    hotwtR4 = 1
    hotwtR5 = 1
    wtR = 1
    wtR2 = 1
    wtR3 = 1
    WWA = 1
    WtrMtr = 1
    WtrMtr2 = 1
    WtrMtr3 = 1
    WtrMtr4 = 1
    WtrMtr5 = 1
    WtrMtr6 = 1
    WtrMtr7 = 1
    WtrMtr8 = 1
    WtrMtr9 = 1
    WtrMtr10 = 1
    WtrMtr11 = 1
    SewerCons = 1
    Sewer2Cons = 1
    Sewer = 1
    Sewer2 = 1
    Base = 1
    SewerBase = 1
    FireSup = 1
    RegComp = 1
    XX = 1
    ThrmCol = 1
    ThrmCnCol = 1
    ThrmAdCol = 1
    ThrmMCol = 1
    BadDtCol = 1
    COVIDRecCol = 1
    MidPeakCons = 1
    OffPeakCons = 1
    OnPeakCons = 1
    MidPeak = 1
    OffPeak = 1
    OnPeak = 1
    ElecMeterCons = 1
    CapCostRecCol = 1
    StormWaterDrainCol = 1
    StormWaterSewerCol = 1
    Cooling_Col = 1
    CoolingCon_Col = 1
    Cooling2_Col = 1
    Cooling2Con_Col = 1
    Cooling3_Col = 1
    Cooling3Con_Col = 1
    Cooling4_Col = 1
    Cooling4Con_Col = 1
    Heating_Col = 1
    HeatingCon_Col = 1
    Heating2_Col = 1
    Heating2Con_Col = 1
    Heating3_Col = 1
    Heating3Con_Col = 1
    Heating4_Col = 1
    Heating4Con_Col = 1
    ThermalMeter2_Col = 1
    ThermalMeter3_Col = 1
    ThermalMeter4_Col = 1
    ThermalMeter5_Col = 1
    ThermalMeter6_Col = 1
    ThermalMeter7_Col = 1
    ThermalMeter8_Col = 1
    Gas_Col = 1
    GasCon_Col = 1
    GasCust_Col = 1
    GasAdmin_Col = 1
    GasMeter_Col = 1
    GasMeter2_Col = 1
    EbillAddie_Col = 1
    PaperBillTax_Col = 1
    HasPyArr_Col = 1
    PAMonthly_Col = 1
    PrevBal_Col = 1
    PastDueArr_Col = 1
    'Call Init_Columns

    ''Assigns a number to each column to referance
    ''Note for Keith: Reorgaize based on charge type?  Can this be done with an Array?
    
    Do Until Cells(1, MIDATE) = "Move In"
        MIDATE = MIDATE + 1
    Loop
    
    Do Until Cells(1, UTILDEP) = "Utility Deposit" Or Cells(1, UTILDEP) = Empty
        UTILDEP = UTILDEP + 1
    Loop
        
    Do Until Cells(1, XX) = Empty
        XX = XX + 1
    Loop
        
    XY = XX + 1
    Cells(1, XX) = "HST"
    Cells(1, XY) = "REBATE"
        
    Do Until Cells(1, AutoPay) = "Enrolled In Auto Pay"
        AutoPay = AutoPay + 1
    Loop
        
    Do Until Cells(1, DepStat) = "Wyse Deposit Status"
        DepStat = DepStat + 1
    Loop
       
    Do Until Cells(1, DepAmt) = "Wyse Deposit Amount"
        DepAmt = DepAmt + 1
    Loop
        
    Do Until Cells(1, AcctSUP) = "Acct. Setup" Or Cells(1, AcctSUP) = Empty
        AcctSUP = AcctSUP + 1
    Loop
    
    Do Until Cells(1, SUPHST) = "HST Account Setup" Or Cells(1, SUPHST) = Empty
        SUPHST = SUPHST + 1
    Loop
    
    Do Until Cells(1, Delivcons) = "Delivery Cons" Or Cells(1, Delivcons) = Empty
        Delivcons = Delivcons + 1
    Loop
        
    Do Until Cells(1, EnCons) = "Wyse Electricity Cons" Or Cells(1, EnCons) = Empty
        EnCons = EnCons + 1
    Loop
    
    Do Until Cells(1, EnCons_2) = "Wyse Electricity 2 Cons" Or Cells(1, EnCons_2) = Empty
        EnCons_2 = EnCons_2 + 1
    Loop
    
    Do Until Cells(1, EnCons_3) = "Wyse Electricity 3 Cons" Or Cells(1, EnCons_3) = Empty
        EnCons_3 = EnCons_3 + 1
    Loop
        
    Do Until Cells(1, LILOcons) = "Line Loss Adjustment Cons" Or Cells(1, LILOcons) = Empty
        LILOcons = LILOcons + 1
    Loop
        
    Do Until Cells(1, RegCons) = "Regulatory Charges Cons" Or Cells(1, RegCons) = Empty
        RegCons = RegCons + 1
    Loop
        
    Do Until Cells(1, Deliv) = "Delivery" Or Cells(1, Deliv) = Empty
        Deliv = Deliv + 1
    Loop
        
    Do Until Cells(1, Energy) = "Wyse Electricity" Or Cells(1, Energy) = Empty
        Energy = Energy + 1
    Loop
    
    Do Until Cells(1, Energy_2) = "Wyse Electricity 2" Or Cells(1, Energy_2) = Empty
        Energy_2 = Energy_2 + 1
    Loop
    
    Do Until Cells(1, Energy_3) = "Wyse Electricity 3" Or Cells(1, Energy_3) = Empty
        Energy_3 = Energy_3 + 1
    Loop
        
    Do Until Cells(1, LILO) = "Line Loss Adjustment" Or Cells(1, LILO) = Empty
        LILO = LILO + 1
    Loop
        
    Do Until Cells(1, Reg) = "Regulatory Charges" Or Cells(1, Reg) = Empty
        Reg = Reg + 1
    Loop
        
    ''Only populates SDF if there is Energy Consumption.  Why?  To differentiate Electric Service from other service fees.  See WWA
    If Cells(1, EnCons) <> Empty Or Cells(1, MidPeak) <> Empty Then
        SDF = 1
        
        Do Until Cells(1, SDF) = "Service Delivery Fee" Or Cells(1, SDF) = Empty
            SDF = SDF + 1
        Loop
        
    ElseIf Cells(1, EnCons) = Empty And Cells(1, MidPeak) = Empty Then
        SDF = 9999
    End If
        
    Do Until Cells(1, SDF2) = "Service Delivery Fee 2" Or Cells(1, SDF2) = Empty
        SDF2 = SDF2 + 1
    Loop
        
    Do Until Cells(1, RegRR) = "Regulatory Administration" Or Cells(1, RegRR) = Empty
        RegRR = RegRR + 1
    Loop
        
    Do Until Cells(1, ElecMeter) = "Electric Meter" Or Cells(1, ElecMeter) = Empty
        ElecMeter = ElecMeter + 1
    Loop
        
    Do Until Cells(1, ElecMeter2) = "Electric Meter 2" Or Cells(1, ElecMeter2) = Empty
        ElecMeter2 = ElecMeter2 + 1
    Loop
    
    Do Until Cells(1, ElecMeter3) = "Electric Meter 3" Or Cells(1, ElecMeter3) = Empty
        ElecMeter3 = ElecMeter3 + 1
    Loop
    
    Do Until Cells(1, ElecMeter4) = "Electric Meter 4" Or Cells(1, ElecMeter4) = Empty
        ElecMeter4 = ElecMeter4 + 1
    Loop
    
    Do Until Cells(1, ElecMeter5) = "Electric Meter 5" Or Cells(1, ElecMeter5) = Empty
        ElecMeter5 = ElecMeter5 + 1
    Loop
        
    Do Until Cells(1, SSSadmin) = "SSS Admin Charge" Or Cells(1, SSSadmin) = Empty
        SSSadmin = SSSadmin + 1
    Loop
        
    Do Until Cells(1, ElecCust) = "Electric Customer Charge" Or Cells(1, ElecCust) = Empty
        ElecCust = ElecCust + 1
    Loop
        
    Do Until Cells(1, HST) = "HST # 832218960 RT0001" Or Cells(1, HST) = Empty
        HST = HST + 1
    Loop
        
    Do Until Cells(1, WyseHST) = "Wyse HST" Or Cells(1, WyseHST) = Empty
        WyseHST = WyseHST + 1
    Loop
    
    Do Until Cells(1, GST) = "GST # 832218960 RT0001" Or Cells(1, GST) = Empty
        GST = GST + 1
    Loop
        
    Do Until Cells(1, WyseHST) = "Wyse GST" Or Cells(1, WyseGST) = Empty
        WyseGST = WyseGST + 1
    Loop
    
    Do Until Cells(1, PST) = "PST #PST-1051-5607" Or Cells(1, PST) = Empty
        PST = PST + 1
    Loop
        
    Do Until Cells(1, WysePST) = "Wyse PST" Or Cells(1, WysePST) = Empty
        WysePST = WysePST + 1
    Loop
        
    Do Until Cells(1, RegRbate) = "Ontario Electricity Rebate" Or Cells(1, RegRbate) = Empty
        RegRbate = RegRbate + 1
    Loop
        
    Do Until Cells(1, wyseRbate) = "Ontario Electricity Rebate - Wys" Or Cells(1, wyseRbate) = Empty
        wyseRbate = wyseRbate + 1
    Loop
        
    Do Until Cells(1, OESP) = "OESP" Or Cells(1, OESP) = Empty
        OESP = OESP + 1
    Loop
        
    Do Until Cells(1, OESPHST) = "OESP HST" Or Cells(1, OESPHST) = Empty
        OESPHST = OESPHST + 1
    Loop
        
    Do Until Cells(1, OESPRbate) = "Ontario Electricity Rebate - OESP" Or Cells(1, OESPRbate) = Empty
        OESPRbate = OESPRbate + 1
    Loop
    
    Do Until Cells(1, coldwtRcons) = "Cold Water Cons" Or Cells(1, coldwtRcons) = Empty
        coldwtRcons = coldwtRcons + 1
    Loop
        
    Do Until Cells(1, coldwtR2cons) = "Cold Water 2 Cons" Or Cells(1, coldwtR2cons) = Empty
        coldwtR2cons = coldwtR2cons + 1
    Loop
    
    Do Until Cells(1, coldwtR3cons) = "Cold Water 3 Cons" Or Cells(1, coldwtR3cons) = Empty
        coldwtR3cons = coldwtR3cons + 1
    Loop
    
    Do Until Cells(1, coldwtR4cons) = "Cold Water 4 Cons" Or Cells(1, coldwtR4cons) = Empty
        coldwtR4cons = coldwtR4cons + 1
    Loop
    
    Do Until Cells(1, coldwtR5cons) = "Cold Water 5 Cons" Or Cells(1, coldwtR5cons) = Empty
        coldwtR5cons = coldwtR5cons + 1
    Loop
        
    Do Until Cells(1, hotwtRcons) = "Hot Water Cons" Or Cells(1, hotwtRcons) = Empty
        hotwtRcons = hotwtRcons + 1
    Loop
    
    Do Until Cells(1, hotwtR2cons) = "Hot Water 2 Cons" Or Cells(1, hotwtR2cons) = Empty
        hotwtR2cons = hotwtR2cons + 1
    Loop
    
    Do Until Cells(1, hotwtR3cons) = "Hot Water 3 Cons" Or Cells(1, hotwtR3cons) = Empty
        hotwtR3cons = hotwtR3cons + 1
    Loop
    
    Do Until Cells(1, hotwtR4cons) = "Hot Water 4 Cons" Or Cells(1, hotwtR4cons) = Empty
        hotwtR4cons = hotwtR4cons + 1
    Loop
    
    Do Until Cells(1, hotwtR5cons) = "Hot Water 5 Cons" Or Cells(1, hotwtR5cons) = Empty
        hotwtR5cons = hotwtR5cons + 1
    Loop
        
    Do Until Cells(1, wtRcons) = "Water Cons" Or Cells(1, wtRcons) = Empty
        wtRcons = wtRcons + 1
    Loop
    
    Do Until Cells(1, wtR2cons) = "Water 2 Cons" Or Cells(1, wtR2cons) = Empty
        wtR2cons = wtR2cons + 1
    Loop
    
    Do Until Cells(1, wtR3cons) = "Water 3 Cons" Or Cells(1, wtR3cons) = Empty
        wtR3cons = wtR3cons + 1
    Loop
        
    Do Until Cells(1, coldwtR) = "Cold Water" Or Cells(1, coldwtR) = Empty
        coldwtR = coldwtR + 1
    Loop
    
    Do Until Cells(1, coldwtR2) = "Cold Water 2" Or Cells(1, coldwtR2) = Empty
        coldwtR2 = coldwtR2 + 1
    Loop
    
    Do Until Cells(1, coldwtR3) = "Cold Water 3" Or Cells(1, coldwtR3) = Empty
        coldwtR3 = coldwtR3 + 1
    Loop
    
    Do Until Cells(1, coldwtR4) = "Cold Water 4" Or Cells(1, coldwtR4) = Empty
        coldwtR4 = coldwtR4 + 1
    Loop
    
    Do Until Cells(1, coldwtR5) = "Cold Water 5" Or Cells(1, coldwtR5) = Empty
        coldwtR5 = coldwtR5 + 1
    Loop
        
    Do Until Cells(1, hotwtR) = "Hot Water" Or Cells(1, hotwtR) = Empty
        hotwtR = hotwtR + 1
    Loop
    
    Do Until Cells(1, hotwtR2) = "Hot Water 2" Or Cells(1, hotwtR2) = Empty
        hotwtR2 = hotwtR2 + 1
    Loop
    
    Do Until Cells(1, hotwtR3) = "Hot Water 3" Or Cells(1, hotwtR3) = Empty
        hotwtR3 = hotwtR3 + 1
    Loop
    
    Do Until Cells(1, hotwtR4) = "Hot Water 4" Or Cells(1, hotwtR4) = Empty
        hotwtR4 = hotwtR4 + 1
    Loop
    
    Do Until Cells(1, hotwtR5) = "Hot Water 5" Or Cells(1, hotwtR5) = Empty
        hotwtR5 = hotwtR5 + 1
    Loop
        
    Do Until Cells(1, wtR) = "Water" Or Cells(1, wtR) = Empty
        wtR = wtR + 1
    Loop
    
    Do Until Cells(1, wtR2) = "Water 2" Or Cells(1, wtR2) = Empty
        wtR2 = wtR2 + 1
    Loop
    
    Do Until Cells(1, wtR3) = "Water 3" Or Cells(1, wtR3) = Empty
        wtR3 = wtR3 + 1
    Loop
        
    Do Until Cells(1, WWA) = "Wyse Water Admin" Or Cells(1, WWA) = Empty
        WWA = WWA + 1
    Loop
        
    If SDF = 9999 Then
        WWA = 1
        Do Until Cells(1, WWA) = "Service Delivery Fee" Or Cells(1, WWA) = "Wyse Water Admin" Or Cells(1, WWA) = Empty
            WWA = WWA + 1
        Loop
    End If
        
    Do Until Cells(1, WtrMtr) = "Water Meter" Or Cells(1, WtrMtr) = Empty
        WtrMtr = WtrMtr + 1
    Loop
    
    Do Until Cells(1, WtrMtr2) = "Water Meter 2" Or Cells(1, WtrMtr2) = Empty
        WtrMtr2 = WtrMtr2 + 1
    Loop
    
    Do Until Cells(1, WtrMtr3) = "Water Meter 3" Or Cells(1, WtrMtr3) = Empty
        WtrMtr3 = WtrMtr3 + 1
    Loop
    
    Do Until Cells(1, WtrMtr4) = "Water Meter 4" Or Cells(1, WtrMtr4) = Empty
        WtrMtr4 = WtrMtr4 + 1
    Loop
    
    Do Until Cells(1, WtrMtr5) = "Water Meter 5" Or Cells(1, WtrMtr5) = Empty
        WtrMtr5 = WtrMtr5 + 1
    Loop
    
    Do Until Cells(1, WtrMtr6) = "Water Meter 6" Or Cells(1, WtrMtr6) = Empty
        WtrMtr6 = WtrMtr6 + 1
    Loop
    
    Do Until Cells(1, WtrMtr7) = "Water Meter 7" Or Cells(1, WtrMtr7) = Empty
        WtrMtr7 = WtrMtr7 + 1
    Loop
    
    Do Until Cells(1, WtrMtr8) = "Water Meter 8" Or Cells(1, WtrMtr8) = Empty
        WtrMtr8 = WtrMtr8 + 1
    Loop
    
    Do Until Cells(1, WtrMtr9) = "Water Meter 9" Or Cells(1, WtrMtr9) = Empty
        WtrMtr9 = WtrMtr9 + 1
    Loop
    
    Do Until Cells(1, WtrMtr10) = "Water Meter 10" Or Cells(1, WtrMtr10) = Empty
        WtrMtr10 = WtrMtr10 + 1
    Loop
    
    Do Until Cells(1, WtrMtr11) = "Water Meter 11" Or Cells(1, WtrMtr11) = Empty
        WtrMtr11 = WtrMtr11 + 1
    Loop
        
    Do Until Cells(1, SewerCons) = "Sewer Cons" Or Cells(1, SewerCons) = Empty
        SewerCons = SewerCons + 1
    Loop
    
    Do Until Cells(1, Sewer2Cons) = "Sewer 2 Cons" Or Cells(1, Sewer2Cons) = Empty
        Sewer2Cons = Sewer2Cons + 1
    Loop
        
    Do Until Cells(1, Sewer) = "Sewer" Or Cells(1, Sewer) = Empty
        Sewer = Sewer + 1
    Loop
    
    Do Until Cells(1, Sewer2) = "Sewer 2" Or Cells(1, Sewer2) = Empty
        Sewer2 = Sewer2 + 1
    Loop
        
    Do Until Cells(1, Base) = "Water Base Charge" Or Cells(1, Base) = Empty
        Base = Base + 1
    Loop
        
    Do Until Cells(1, SewerBase) = "Sewer Base" Or Cells(1, SewerBase) = Empty
        SewerBase = SewerBase + 1
    Loop
        
    Do Until Cells(1, FireSup) = "Fire Supply" Or Cells(1, FireSup) = Empty
        FireSup = FireSup + 1
    Loop
    
    Do Until Cells(1, RegComp) = "Regulatory Assessment" Or Cells(1, RegComp) = Empty
        RegComp = RegComp + 1
    Loop
        
    Do Until Cells(1, OESPRbateCR) = "Ontario Electricity Rebate - OESP Credit -" Or Cells(1, OESPRbateCR) = Empty
        OESPRbateCR = OESPRbateCR + 1
    Loop
        
    Do Until Cells(1, OESPHSTCR) = "OESP HST Credit -" Or Cells(1, OESPHSTCR) = Empty
        OESPHSTCR = OESPHSTCR + 1
    Loop
        
    Do Until Cells(1, ThrmCol) = "Thermal Charge" Or Cells(1, ThrmCol) = Empty
        ThrmCol = ThrmCol + 1
    Loop
        
    Do Until Cells(1, ThrmMCol) = "Thermal Meter" Or Cells(1, ThrmMCol) = Empty
        ThrmMCol = ThrmMCol + 1
    Loop
    
    Do Until Cells(1, ThermalMeter2_Col) = "Thermal Meter 2" Or Cells(1, ThermalMeter2_Col) = Empty
        ThermalMeter2_Col = ThermalMeter2_Col + 1
    Loop
        
    Do Until Cells(1, ThermalMeter3_Col) = "Thermal Meter 3" Or Cells(1, ThermalMeter3_Col) = Empty
        ThermalMeter3_Col = ThermalMeter3_Col + 1
    Loop
    
    Do Until Cells(1, ThermalMeter4_Col) = "Thermal Meter 4" Or Cells(1, ThermalMeter4_Col) = Empty
        ThermalMeter4_Col = ThermalMeter4_Col + 1
    Loop
    
    Do Until Cells(1, ThermalMeter5_Col) = "Thermal Meter 5" Or Cells(1, ThermalMeter5_Col) = Empty
        ThermalMeter5_Col = ThermalMeter5_Col + 1
    Loop
    
    Do Until Cells(1, ThermalMeter6_Col) = "Thermal Meter 6" Or Cells(1, ThermalMeter6_Col) = Empty
        ThermalMeter6_Col = ThermalMeter6_Col + 1
    Loop
    
    Do Until Cells(1, ThermalMeter7_Col) = "Thermal Meter 7" Or Cells(1, ThermalMeter7_Col) = Empty
        ThermalMeter7_Col = ThermalMeter7_Col + 1
    Loop
    
    Do Until Cells(1, ThermalMeter8_Col) = "Thermal Meter 8" Or Cells(1, ThermalMeter8_Col) = Empty
        ThermalMeter8_Col = ThermalMeter8_Col + 1
    Loop
        
    Do Until Cells(1, ThrmAdCol) = "Thermal Admin" Or Cells(1, ThrmAdCol) = Empty
        ThrmAdCol = ThrmAdCol + 1
    Loop
        
    Do Until Cells(1, ThrmCnCol) = "Thermal Charge Cons" Or Cells(1, ThrmCnCol) = Empty
        ThrmCnCol = ThrmCnCol + 1
    Loop
        
    Do Until Cells(1, BadDtCol) = "Bad Debt Recovery" Or Cells(1, BadDtCol) = Empty
        BadDtCol = BadDtCol + 1
    Loop
        
    Do Until Cells(1, COVIDRecCol) = "COVID Recovery Fee" Or Cells(1, COVIDRecCol) = Empty
        COVIDRecCol = COVIDRecCol + 1
    Loop
        
    Do Until Cells(1, MidPeakCons) = "Mid-Peak Cons" Or Cells(1, MidPeakCons) = Empty
        MidPeakCons = MidPeakCons + 1
    Loop
        
    Do Until Cells(1, OffPeakCons) = "Off-Peak Cons" Or Cells(1, OffPeakCons) = Empty
        OffPeakCons = OffPeakCons + 1
    Loop
        
    Do Until Cells(1, OnPeakCons) = "On-Peak Cons" Or Cells(1, OnPeakCons) = Empty
        OnPeakCons = OnPeakCons + 1
    Loop
    
    Do Until Cells(1, MidPeak) = "Mid-Peak" Or Cells(1, MidPeak) = Empty
        MidPeak = MidPeak + 1
    Loop
        
    Do Until Cells(1, OffPeak) = "Off-Peak" Or Cells(1, OffPeak) = Empty
        OffPeak = OffPeak + 1
    Loop
        
    Do Until Cells(1, OnPeak) = "On-Peak" Or Cells(1, OnPeak) = Empty
        OnPeak = OnPeak + 1
    Loop
    
    Do Until Cells(1, ElecMeterCons) = "Electric Meter Cons" Or Cells(1, ElecMeterCons) = Empty
        ElecMeterCons = ElecMeterCons + 1
    Loop
    
    Do Until Cells(1, CapCostRecCol) = "Capital Cost Recovery" Or Cells(1, CapCostRecCol) = Empty
        CapCostRecCol = CapCostRecCol + 1
    Loop
    
    Do Until Cells(1, StormWaterDrainCol) = "Storm Water Drainage" Or Cells(1, StormWaterDrainCol) = Empty
        StormWaterDrainCol = StormWaterDrainCol + 1
    Loop
    
    Do Until Cells(1, StormWaterSewerCol) = "Storm Sewer" Or Cells(1, StormWaterSewerCol) = Empty
        StormWaterSewerCol = StormWaterSewerCol + 1
    Loop
    
    Do Until Cells(1, Cooling_Col) = "Cooling Charge" Or Cells(1, Cooling_Col) = Empty
        Cooling_Col = Cooling_Col + 1
    Loop
    
    Do Until Cells(1, CoolingCon_Col) = "Cooling Charge Cons" Or Cells(1, CoolingCon_Col) = Empty
        CoolingCon_Col = CoolingCon_Col + 1
    Loop
    
    Do Until Cells(1, Cooling2_Col) = "Cooling Charge 2" Or Cells(1, Cooling2_Col) = Empty
        Cooling2_Col = Cooling2_Col + 1
    Loop
    
    Do Until Cells(1, Cooling2Con_Col) = "Cooling Charge 2 Cons" Or Cells(1, Cooling2Con_Col) = Empty
        Cooling2Con_Col = Cooling2Con_Col + 1
    Loop
    
    Do Until Cells(1, Cooling3_Col) = "Cooling Charge 3" Or Cells(1, Cooling3_Col) = Empty
        Cooling3_Col = Cooling3_Col + 1
    Loop
    
    Do Until Cells(1, Cooling3Con_Col) = "Cooling Charge 3 Cons" Or Cells(1, Cooling3Con_Col) = Empty
        Cooling3Con_Col = Cooling3Con_Col + 1
    Loop
    
    Do Until Cells(1, Cooling4_Col) = "Cooling Charge 4" Or Cells(1, Cooling4_Col) = Empty
        Cooling4_Col = Cooling4_Col + 1
    Loop
    
    Do Until Cells(1, Cooling4Con_Col) = "Cooling Charge 4 Cons" Or Cells(1, Cooling4Con_Col) = Empty
        Cooling4Con_Col = Cooling4Con_Col + 1
    Loop
        
    Do Until Cells(1, Heating_Col) = "Heating Charge" Or Cells(1, Heating_Col) = Empty
        Heating_Col = Heating_Col + 1
    Loop
    
    Do Until Cells(1, HeatingCon_Col) = "Heating Charge Cons" Or Cells(1, HeatingCon_Col) = Empty
        HeatingCon_Col = HeatingCon_Col + 1
    Loop
    
    Do Until Cells(1, Heating2_Col) = "Heating Charge 2" Or Cells(1, Heating2_Col) = Empty
        Heating2_Col = Heating2_Col + 1
    Loop
    
    Do Until Cells(1, Heating2Con_Col) = "Heating Charge 2 Cons" Or Cells(1, Heating2Con_Col) = Empty
        Heating2Con_Col = Heating2Con_Col + 1
    Loop
    
    Do Until Cells(1, Heating3_Col) = "Heating Charge 3" Or Cells(1, Heating3_Col) = Empty
        Heating3_Col = Heating3_Col + 1
    Loop
    
    Do Until Cells(1, Heating3Con_Col) = "Heating Charge 3 Cons" Or Cells(1, Heating3Con_Col) = Empty
        Heating3Con_Col = Heating3Con_Col + 1
    Loop
    
    Do Until Cells(1, Heating4_Col) = "Heating Charge 4" Or Cells(1, Heating4_Col) = Empty
        Heating4_Col = Heating4_Col + 1
    Loop
    
    Do Until Cells(1, Heating4Con_Col) = "Heating Charge 4 Cons" Or Cells(1, Heating4Con_Col) = Empty
        Heating4Con_Col = Heating4Con_Col + 1
    Loop
    
    Do Until Cells(1, Gas_Col) = "Gas" Or Cells(1, Gas_Col) = Empty
        Gas_Col = Gas_Col + 1
    Loop
    
    Do Until Cells(1, GasCon_Col) = "Gas Cons" Or Cells(1, GasCon_Col) = Empty
        GasCon_Col = GasCon_Col + 1
    Loop
    
    Do Until Cells(1, GasCust_Col) = "Customer Charge" Or Cells(1, GasCust_Col) = "Gas Customer Charge" Or Cells(1, GasCust_Col) = Empty
        GasCust_Col = GasCust_Col + 1
    Loop
    
    Do Until Cells(1, GasAdmin_Col) = "Wyse Gas Admin" Or Cells(1, GasAdmin_Col) = "Gas Admin Charge" Or Cells(1, GasAdmin_Col) = Empty
        GasAdmin_Col = GasAdmin_Col + 1
    Loop
    
    Do Until Cells(1, GasMeter_Col) = "Gas Meter" Or Cells(1, GasMeter_Col) = Empty
        GasMeter_Col = GasMeter_Col + 1
    Loop
    
    Do Until Cells(1, GasMeter2_Col) = "Gas Meter 2" Or Cells(1, GasMeter2_Col) = Empty
        GasMeter2_Col = GasMeter2_Col + 1
    Loop
    
    Do Until Cells(1, EbillAddie_Col) = "E-Bill Email" Or Cells(1, EbillAddie_Col) = Empty
        EbillAddie_Col = EbillAddie_Col + 1
    Loop
    
    Do Until Cells(1, PaperBillTax_Col) = "Paper Bill Fee" Or Cells(1, PaperBillTax_Col) = Empty
        PaperBillTax_Col = PaperBillTax_Col + 1
    Loop
    
    Do Until Cells(1, HasPyArr_Col) = "Has Payment Arrangement" Or Cells(1, HasPyArr_Col) = Empty
        HasPyArr_Col = HasPyArr_Col + 1
    Loop
    
    Do Until Cells(1, PAMonthly_Col) = "Monthly Arranged Amount" Or Cells(1, PAMonthly_Col) = Empty
        PAMonthly_Col = PAMonthly_Col + 1
    Loop

    Do Until Cells(1, PrevBal_Col) = "Previous Balance" Or Cells(1, PrevBal_Col) = Empty
        PrevBal_Col = PrevBal_Col + 1
    Loop
    
    Do Until Cells(1, PastDueArr_Col) = "Past Due Arrangement" Or Cells(1, PastDueArr_Col) = Empty
        PastDueArr_Col = PastDueArr_Col + 1
    Loop
        
    '' get info from spreadsheet

    'Taxes
    If Cells(1, HST) <> Empty Or Cells(1, WyseHST) <> Empty Then
        HST_Rate = Find_info("HST", 2, PropSS, PropSSshort, HST)
    End If
    
    If Cells(1, GST) <> Empty Or Cells(1, WyseGST) <> Empty Then
        GST_Rate = Find_info("GST", 2, PropSS, PropSSshort, GST)
    End If
    
    If Cells(1, PST) <> Empty Or Cells(1, WysePST) <> Empty Then
        PST_Rate = Find_info("PST", 2, PropSS, PropSSshort, PST)
    End If
    
    If answr2 = 6 Then
        If Cells(1, RegRbate) <> Empty Or Cells(1, wyseRbate) <> Empty Then
            OER_Rate = Find_info("OER", 2, PropSS, PropSSshort, RegRbate)
        End If
    End If
        
    ''Electric Flat Fees
    If Cells(1, Delivcons) <> Empty Then
        LAF = Find_info("LAF", 2, PropSS, PropSSshort)
    End If
        
    If Cells(1, Reg) <> Empty Then
        RegChg_Rate = Find_rate("Regulatory_Charge", PropSS, PropSSshort, Reg) 'New Code
    End If
    
    If Cells(1, SSSadmin) <> Empty Then
        SSS_Admin = Find_info("SSS_Admin", 3, PropSS, PropSSshort, SSSadmin)
    End If
    
    If Cells(1, RegRR) <> Empty Then
        Regulatory_Admin = Find_info("Regulatory_Admin", 3, PropSS, PropSSshort, RegRR)
    End If
    
    If Cells(1, ElecMeter) <> Empty Then
        If Cells(1, ElecMeterCons) = Empty Then
            Electric_Meter = Find_info("Electric_Meter", 3, PropSS, PropSSshort, ElecMeter) 'flat fee
        ElseIf Cells(1, ElecMeterCons) <> Empty Then
            ElecMeter_Rate = Find_rate("Electric_Meter", PropSS, PropSSshort, ElecMeter) 'Metered; New Code
        End If
    End If
    
    If Cells(1, ElecMeter2) <> Empty Then
        Electric_Meter2 = Find_info("Electric_Meter2", 3, PropSS, PropSSshort, ElecMeter2)
    End If
    
    If Cells(1, ElecMeter3) <> Empty Then
        Electric_Meter3 = Find_info("Electric_Meter3", 3, PropSS, PropSSshort, ElecMeter3)
    End If
    
    If Cells(1, ElecMeter4) <> Empty Then
        Electric_Meter4 = Find_info("Electric_Meter4", 3, PropSS, PropSSshort, ElecMeter4)
    End If
    
    If Cells(1, ElecMeter5) <> Empty Then
        Electric_Meter5 = Find_info("Electric_Meter5", 3, PropSS, PropSSshort, ElecMeter5)
    End If
    
    If Cells(1, SDF) <> Empty Then
        Electric_Service_Delivery = Find_info("Electric_Service_Delivery", 3, PropSS, PropSSshort, SDF)
    End If
    
    If Cells(1, Deliv) <> Empty Then
           Delivery_Rate = Find_rate("Delivery_Charge", PropSS, PropSSshort, Deliv) ' New Code
    End If
    
    If Cells(1, Energy) <> Empty Then
        Energy_Rate() = Find_rate("Energy_Rate", PropSS, PropSSshort, Energy) ' New Code
    End If
    
    If Cells(1, Energy_2) <> Empty Then
        Energy2_Rate() = Find_rate("Energy2_Rate", PropSS, PropSSshort, Energy_2) ' New Code
        If Array_Compare(Energy_Rate, Energy2_Rate) Then
            E1E2_MM_Answer = MsgBox("Electricity 1 & 2 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
    End If

    If Cells(1, Energy_3) <> Empty Then
        Energy3_Rate() = Find_rate("Energy3_Rate", PropSS, PropSSshort, Energy_3) ' New Code
        If Array_Compare(Energy_Rate, Energy3_Rate) Then
            E1E3_MM_Answer = MsgBox("Electricity 1 & 3 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
        
        If Array_Compare(Energy2_Rate, Energy3_Rate) Then
            E2E3_MM_Answer = MsgBox("Electricity 2 & 3 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
    End If
    
    ''''TOU Energy Rates
    If Cells(1, MidPeakCons) <> Empty And Cells(1, OnPeakCons) <> Empty And Cells(1, OffPeakCons) <> Empty Then
        If Cells(1, MidPeak) <> Empty Then
            MidPeak_Rate() = Find_rate("Mid_Peak", PropSS, PropSSshort, MidPeak) 'New Code
        End If
        
        If Cells(1, OffPeak) <> Empty Then
            OffPeak_Rate() = Find_rate("Off_Peak", PropSS, PropSSshort, OffPeak) 'New Code
        End If
        
        If Cells(1, OnPeak) <> Empty Then
            OnPeak_Rate() = Find_rate("On_Peak", PropSS, PropSSshort, OnPeak) 'New Code
        End If
    End If
    
    ''water
    If Cells(1, wtR) <> Empty Then
        Water_Rate() = Find_rate("Water_Rate", PropSS, PropSSshort, wtR) 'New Code
    End If
    
    If Cells(1, wtR2) <> Empty Then
        Water2_Rate() = Find_rate("Water2_Rate", PropSS, PropSSshort, wtR) 'New Code
        If Array_Compare(Water_Rate, Water2_Rate) Then
            W1W2_MM_Answer = MsgBox("Water 1 and 2 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
    End If
    
    If Cells(1, wtR3) <> Empty Then
        Water3_Rate() = Find_rate("Water3_Rate", PropSS, PropSSshort, wtR) 'New Code
        If Array_Compare(Water_Rate, Water3_Rate) Then
            W1W3_MM_Answer = MsgBox("Water 1 & 3 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
        
        If Array_Compare(Water2_Rate, Water3_Rate) Then
            W2W3_MM_Answer = MsgBox("Water 2 & 3 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
    End If
    
    If Cells(1, coldwtR) <> Empty Then
        clDWater_Rate() = Find_rate("Water_Rate", PropSS, PropSSshort, coldwtR) 'New Code
    End If
    
    If Cells(1, hotwtR) <> Empty Then
        hTWater_Rate() = Find_rate("Water_Rate", PropSS, PropSSshort, hotwtR) 'New Code
        If Array_Compare(clDWater_Rate, hTWater_Rate) Then
            CW1HW1_MM_Answer = MsgBox("Cold Water 1 and Hot Water 1 Rates Match.  Is this a Multi-meter property?", vbYesNo)
        End If
    End If
    
    If Cells(1, coldwtR2) <> Empty Then
        clDWater2_Rate() = Find_rate("Water2_Rate", PropSS, PropSSshort, coldwtR2) 'New Code
    End If
    
    If Cells(1, coldwtR3) <> Empty Then
        clDWater3_Rate() = Find_rate("Water3_Rate", PropSS, PropSSshort, coldwtR3) 'New Code
    End If
    
    If Cells(1, coldwtR4) <> Empty Then
        clDWater4_Rate() = Find_rate("Water4_Rate", PropSS, PropSSshort, coldwtR4) 'New Code
    End If
    
    If Cells(1, coldwtR5) <> Empty Then
        clDWater5_Rate() = Find_rate("Water5_Rate", PropSS, PropSSshort, coldwtR5) 'New Code
    End If
    
    If Cells(1, hotwtR2) <> Empty Then
        hTWater2_Rate() = Find_rate("Water2_Rate", PropSS, PropSSshort, hotwtR2) 'New Code
    End If
    
    If Cells(1, hotwtR3) <> Empty Then
        hTWater3_Rate() = Find_rate("Water3_Rate", PropSS, PropSSshort, hotwtR3) 'New Code
    End If
    
    If Cells(1, hotwtR4) <> Empty Then
        hTWater4_Rate() = Find_rate("Water4_Rate", PropSS, PropSSshort, hotwtR4) 'New Code
    End If
    
    If Cells(1, hotwtR5) <> Empty Then
        hTWater5_Rate() = Find_rate("Water5_Rate", PropSS, PropSSshort, hotwtR5) 'New Code
    End If
    
    If Cells(1, WWA) <> Empty Then
        Water_Service = Find_info("Water_Service", 3, PropSS, PropSSshort, WWA)
    End If
    
    If Cells(1, WtrMtr) <> Empty Then
        Water_Meter = Find_info("Water_Meter", 3, PropSS, PropSSshort, WtrMtr)
    End If
    
    If Cells(1, WtrMtr2) <> Empty Then
        Water_Meter_2 = Find_info("Water_Meter_2", 3, PropSS, PropSSshort, WtrMtr2)
    End If
    
    If Cells(1, WtrMtr3) <> Empty Then
        Water_Meter_3 = Find_info("Water_Meter_3", 3, PropSS, PropSSshort, WtrMtr3)
    End If
    
    If Cells(1, WtrMtr4) <> Empty Then
        Water_Meter_4 = Find_info("Water_Meter_4", 3, PropSS, PropSSshort, WtrMtr4)
    End If
    
    If Cells(1, WtrMtr5) <> Empty Then
        Water_Meter_5 = Find_info("Water_Meter_5", 3, PropSS, PropSSshort, WtrMtr5)
    End If
    
    If Cells(1, WtrMtr6) <> Empty Then
        Water_Meter_6 = Find_info("Water_Meter_6", 3, PropSS, PropSSshort, WtrMtr6)
    End If
    
    If Cells(1, WtrMtr7) <> Empty Then
        Water_Meter_7 = Find_info("Water_Meter_7", 3, PropSS, PropSSshort, WtrMtr7)
    End If
    
    If Cells(1, WtrMtr8) <> Empty Then
        Water_Meter_8 = Find_info("Water_Meter_8", 3, PropSS, PropSSshort, WtrMtr8)
    End If
    
    If Cells(1, WtrMtr9) <> Empty Then
        Water_Meter_9 = Find_info("Water_Meter_9", 3, PropSS, PropSSshort, WtrMtr9)
    End If
    
    If Cells(1, WtrMtr10) <> Empty Then
        Water_Meter_10 = Find_info("Water_Meter_10", 3, PropSS, PropSSshort, WtrMtr10)
    End If
    
    If Cells(1, WtrMtr11) <> Empty Then
        Water_Meter_11 = Find_info("Water_Meter_11", 3, PropSS, PropSSshort, WtrMtr11)
    End If
    
    ''Miscellaneous charges; reorganize by utility/admin
    If Cells(1, SDF2) <> Empty Then
        Electric_Service_Delivery_2 = Find_info("Electric_Service_Delivery_2", 3, PropSS, PropSSshort, SDF2)
    End If
    
    If Cells(1, Sewer) <> Empty Then
        Sewer_Rate() = Find_rate("Sewer_Rate", PropSS, PropSSshort, Sewer) 'New Code
    End If
    
    If Cells(1, Sewer2) <> Empty Then
        Sewer2_Rate() = Find_rate("Sewer2_Rate", PropSS, PropSSshort, Sewer) 'New Code
'        If Array_Compare(Sewer_Rate, Sewer2_Rate) Then
'            Sew_MM_Answer = MsgBox("Rates Match.  Is this a Multi-meter property?", vbYesNo)
'        Else
'            Sew_MM_Answer = MsgBox("Rates are unique.", vbOKOnly)
'        End If
    End If
    
    If Cells(1, RegComp) <> Empty Then
        Regulatory_Assessment = Find_info("Regulatory_Assessment", 3, PropSS, PropSSshort, RegComp)
    End If
    
    If Cells(1, RegRR) <> Empty Then
        Regulatory_Admin = Find_info("Regulatory_Admin", 3, PropSS, PropSSshort, RegRR)
    End If
    
    If Cells(1, BadDtCol) <> Empty Then
        Bad_Debt = Find_info("Bad_Debt", 3, PropSS, PropSSshort, BadDtCol)
    End If
    
    If Cells(1, ThrmCol) <> Empty Then
        Thermal_Rate() = Find_rate("Thermal_Rate", PropSS, PropSSshort, ThrmCol) 'New Code
    End If
    
    If Cells(1, ThrmAdCol) <> Empty Then
        Thermal_Admin = Find_info("Thermal_Admin", 3, PropSS, PropSSshort, ThrmAdCol)
    End If
    
    If Cells(1, ThrmMCol) <> Empty Then
        Thermal_Meter = Find_info("Thermal_Meter", 3, PropSS, PropSSshort, ThrmMCol)
    End If
    
    If Cells(1, ThermalMeter2_Col) <> Empty Then
        Thermal_Meter_2 = Find_info("Thermal_Meter_2", 3, PropSS, PropSSshort, ThermalMeter2_Col)
    End If
    
    If Cells(1, ThermalMeter3_Col) <> Empty Then
        Thermal_Meter_3 = Find_info("Thermal_Meter_3", 3, PropSS, PropSSshort, ThermalMeter3_Col)
    End If
    
    If Cells(1, ThermalMeter4_Col) <> Empty Then
        Thermal_Meter_4 = Find_info("Thermal_Meter_4", 3, PropSS, PropSSshort, ThermalMeter4_Col)
    End If
    
    If Cells(1, ThermalMeter5_Col) <> Empty Then
        Thermal_Meter_5 = Find_info("Thermal_Meter_5", 3, PropSS, PropSSshort, ThermalMeter5_Col)
    End If
    
    If Cells(1, ThermalMeter6_Col) <> Empty Then
        Thermal_Meter_6 = Find_info("Thermal_Meter_6", 3, PropSS, PropSSshort, ThermalMeter6_Col)
    End If
    
    If Cells(1, ThermalMeter7_Col) <> Empty Then
        Thermal_Meter_7 = Find_info("Thermal_Meter_7", 3, PropSS, PropSSshort, ThermalMeter7_Col)
    End If
    
    If Cells(1, ThermalMeter8_Col) <> Empty Then
        Thermal_Meter_8 = Find_info("Thermal_Meter_8", 3, PropSS, PropSSshort, ThermalMeter8_Col)
    End If
    
    If Cells(1, COVIDRecCol) <> Empty Then
        COVID_Rec = Find_info("COVID_Fee", 3, PropSS, PropSSshort, COVIDRecCol)
    End If
    
    If Cells(1, StormWaterSewerCol) <> Empty Then
        StrmWtrSewer = Find_info("Stormwater_Base", 3, PropSS, PropSSshort, StormWaterSewerCol)
    End If
    
    If Cells(1, Cooling_Col) <> Empty Or Cells(1, Cooling2_Col) <> Empty Or Cells(1, Cooling3_Col) <> Empty Or Cells(1, Cooling4_Col) <> Empty Then
        Cooling_Rate() = Find_rate("Cooling_Rate", PropSS, PropSSshort, Cooling_Col) ' New Code
    End If
    
    If Cells(1, Heating_Col) <> Empty Or Cells(1, Heating2_Col) <> Empty Or Cells(1, Heating3_Col) <> Empty Or Cells(1, Heating4_Col) <> Empty Then
        Heating_Rate() = Find_rate("Heating_Rate", PropSS, PropSSshort, Heating_Col) ' New Code
    End If
    
    If Cells(1, Gas_Col) <> Empty Then
        Gas_Rate() = Find_rate("Gas_Rate", PropSS, PropSSshort, Gas_Col) ' New Code
    End If
    
    If Cells(1, GasAdmin_Col) <> Empty Then
        gas_admin = Find_info("Gas_Admin", 3, PropSS, PropSSshort, GasAdmin_Col)
    End If
    
    If Cells(1, GasMeter_Col) <> Empty Then
        gas_meter = Find_info("Gas_Meter", 3, PropSS, PropSSshort, GasMeter_Col)
    End If
    
    If Cells(1, GasMeter2_Col) <> Empty Then
        gas_meter2 = Find_info("Gas_Meter_2", 3, PropSS, PropSSshort, GasMeter2_Col)
    End If
    
    'RUBS
    If Cells(1, ElecCust) <> Empty Then
        elec_cust = FindRUBs_info("Elec_cust", PropSS, PropSSshort, ElecCust)
    End If
    
    If Cells(1, Base) <> Empty Then
        Water_Base = FindRUBs_info("Water_Base", PropSS, PropSSshort, Base)
    End If
        
    If Cells(1, SewerBase) <> Empty Then
        Sewer_Base = FindRUBs_info("Sewer_Base", PropSS, PropSSshort, SewerBase)
    End If
    
    If Cells(1, FireSup) <> Empty Then
        Fire_Supply = FindRUBs_info("Fire_Supply", PropSS, PropSSshort, FireSup)
    End If
    
    If Cells(1, CapCostRecCol) <> Empty Then
        Cap_Cst_Rec = FindRUBs_info("Capital_Cost_Recovery", PropSS, PropSSshort, CapCostRecCol)
    End If
    
    If Cells(1, StormWaterDrainCol) <> Empty Then
        StrmWtrDrain = FindRUBs_info("Stormwater_Charge", PropSS, PropSSshort, StormWaterDrainCol)
    End If
    
    If Cells(1, GasCust_Col) <> Empty Then
        gas_cust = FindRUBs_info("Gas_Cust", PropSS, PropSSshort, GasCust_Col)
    End If
   
    ''Main Loop
    Do Until Cells(QCCurRow, 1) = Empty
        'initalize tax check accumulators on each pass
        HSTAccumulator = 0
        WyseHSTAccumulator = 0
        OERAccumulator = 0
        WyseOERAccumulator = 0
        GSTAccumulator = 0
        WyseGSTAccumulator = 0
        PSTAccumulator = 0
        WysePSTAccumulator = 0
        
        'initialize proration Math
        CycleLength = CycleEndDt - NewDt + 1
        ProLength = CycleEndDt - Cells(QCCurRow, MIDATE).Value + 1
        If ProLength > CycleLength Then
            ProLength = CycleLength
        End If

        If Cells(QCCurRow, MIDATE).Value >= MI_DATE Then
            'Highlight NMI
            Rows(QCCurRow).EntireRow.Select
            tintNMI
            
            'Checks NMI's for utility deposit
            If Cells(QCCurRow, UTILDEP) = Empty Then
                Cells(QCCurRow, UTILDEP).Select
                turnRed
            End If
            
            'Checks NMI's Deposit correct
            ''''Move out of NMI checks
            If Cells(QCCurRow, UTILDEP) = Cells(QCCurRow, DepAmt) Then
                Cells(QCCurRow, UTILDEP).Select
                turnGreen
            End If
            
            'Checks NMI's Account Set up Exists
            If Cells(QCCurRow, AcctSUP) = Empty Then
                Cells(QCCurRow, AcctSUP).Select
                turnRed
            End If

            'Checks NMI's Setup HST Exists
            If Cells(QCCurRow, SUPHST) = Empty Then
                Cells(QCCurRow, SUPHST).Select
                turnRed
            End If
            
            If Cells(QCCurRow, MIDATE) < CycleEndDt And Cells(QCCurRow, Energy) = 0 And Cells(1, Energy) <> Empty Then
                Cells(QCCurRow, Energy).Select
                turnRed
            End If
            If Cells(QCCurRow, MIDATE) < CycleEndDt And Cells(QCCurRow, LILO) = 0 And Cells(1, LILO) <> Empty Then
                Cells(QCCurRow, LILO).Select
                turnRed
            End If
            If Cells(QCCurRow, MIDATE) < CycleEndDt And Cells(QCCurRow, RegComp) = 0 And Cells(1, RegComp) <> Empty Then
                Cells(QCCurRow, RegComp).Select
                turnRed
            End If
            If Cells(QCCurRow, MIDATE) < CycleEndDt And Cells(QCCurRow, Reg) = 0 And Cells(1, Reg) <> Empty Then
                Cells(QCCurRow, Reg).Select
                turnRed
            End If
            If Cells(QCCurRow, MIDATE) < CycleEndDt And Cells(QCCurRow, Deliv) = 0 And Cells(1, Deliv) <> Empty Then
                Cells(QCCurRow, Deliv).Select
                turnRed
            End If
        End If
        
        'LAF Checker
        If Cells(QCCurRow, Delivcons) <> "" And Cells(QCCurRow, EnCons) <> "" And Cells(QCCurRow, RegCons) <> "" Then
            L = Round((Cells(QCCurRow, EnCons) + Cells(QCCurRow, EnCons_2) + Cells(QCCurRow, EnCons_3)) * LAF, 2)
            M = Round((Cells(QCCurRow, EnCons) + Cells(QCCurRow, EnCons_2) + Cells(QCCurRow, EnCons_3)) * (LAF - 1), 2)
            Cells(QCCurRow, XY).Value = "Check"
            If Cells(QCCurRow, Delivcons) = L Then
                Cells(QCCurRow, Delivcons).Select
                turnGreen
            ElseIf Cells(QCCurRow, Delivcons) >= L - 0.01 And Cells(QCCurRow, Delivcons) <= L + 0.01 Then
                Cells(QCCurRow, Delivcons).Select
                tintGreen
            ElseIf Cells(QCCurRow, Delivcons) >= L - 0.02 And Cells(QCCurRow, Delivcons) <= L + 0.02 Then
                Cells(QCCurRow, Delivcons).Select
                warnOrange
            Else
                Cells(QCCurRow, Delivcons).Select
                turnRed
            End If
            
            If Cells(QCCurRow, RegCons) = L Then
                Cells(QCCurRow, RegCons).Select
               turnGreen
            ElseIf Cells(QCCurRow, RegCons) >= L - 0.01 And Cells(QCCurRow, RegCons) <= L + 0.01 Then
                Cells(QCCurRow, RegCons).Select
                tintGreen
            ElseIf Cells(QCCurRow, RegCons) >= L - 0.02 And Cells(QCCurRow, RegCons) <= L + 0.02 Then
                Cells(QCCurRow, RegCons).Select
                warnOrange
            Else
                Cells(QCCurRow, RegCons).Select
                turnRed
            End If
            
            If Cells(1, LILOcons) <> Empty Then
                If Cells(QCCurRow, LILOcons) = M Then
                    Cells(QCCurRow, LILOcons).Select
                    turnGreen
                ElseIf Cells(QCCurRow, LILOcons) >= L - 0.01 And Cells(QCCurRow, LILOcons) <= L + 0.01 Then
                    Cells(QCCurRow, LILOcons).Select
                    tintGreen
                ElseIf Cells(QCCurRow, LILOconss) >= L - 0.02 And Cells(QCCurRow, LILOcons) <= L + 0.02 Then
                    Cells(QCCurRow, LILOcons).Select
                    warnOrange
                Else
                    Cells(QCCurRow, LILOcons).Select
                    turnRed
                End If
            End If
        End If
        
        If Cells(QCCurRow, DepStat) = "Waived" Or Cells(QCCurRow, DepStat) = "Refunded" Then
            If Cells(QCCurRow, DepAmt) = "0" Then
                Cells(QCCurRow, DepAmt).Select
                turnGreen
            Else
                Cells(QCCurRow, DepAmt).Select
                turnRed
            End If
        End If
        
        'Checks for PMT plans that need to be verified
        If Cells(QCCurRow, DepStat) = "Payment Plan" Then
            Cells(QCCurRow, DepStat).Select
            turnRed
            If Cells(QCCurRow, AutoPay) = "No" And Cells(QCCurRow, DepAmt) = Dephigh Then
                Cells(QCCurRow, DepAmt).Select
                turnGreen
            ElseIf Cells(QCCurRow, AutoPay) = "Yes" And Cells(QCCurRow, DepAmt) = DepLow Then
                Cells(QCCurRow, DepAmt).Select
                turnGreen
            Else
                Cells(QCCurRow, DepAmt).Select
                turnRed
            End If
        End If
        
        If Cells(QCCurRow, DepStat) = "New" Or Cells(QCCurRow, DepStat) = "Assessed" Or Cells(QCCurRow, DepStat) = "Paid" Then
            If Cells(QCCurRow, AutoPay) = "No" And Cells(QCCurRow, DepAmt) = Dephigh Then
                Cells(QCCurRow, DepAmt).Select
                turnGreen
            ElseIf Cells(QCCurRow, AutoPay) = "Yes" And Cells(QCCurRow, DepAmt) = DepLow Then
                Cells(QCCurRow, DepAmt).Select
                turnGreen
            Else
                Cells(QCCurRow, DepAmt).Select
                turnRed
            End If
        End If
        
        If Cells(QCCurRow, SUPHST) <> Empty Then
            If Cells(QCCurRow, SUPHST) = Setup_Fee * HST_Rate Then
                Cells(QCCurRow, SUPHST).Select
                turnGreen
            Else
                Cells(QCCurRow, SUPHST).Select
                turnRed
            End If
        End If
        
        If Cells(QCCurRow, AcctSUP) <> Empty Then
            If Cells(QCCurRow, AcctSUP) = Setup_Fee Then
                Cells(QCCurRow, AcctSUP).Select
                turnGreen
            Else
                Cells(QCCurRow, AcctSUP).Select
                turnRed
            End If
        End If
    
        'Per Unit Flat Fees
        If Cells(QCCurRow, SSSadmin) <> Empty Then
            Call checkAndFormat(QCCurRow, SSSadmin, SSS_Admin, CycleLength, ProLength)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, SSSadmin), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, SSSadmin), 2)
        End If
        
        If Cells(1, RegComp) <> Empty Then
            Call checkAndFormat(QCCurRow, RegComp, Regulatory_Assessment, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, RegComp), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, RegComp), 2)
        End If
        
        If Cells(QCCurRow, RegRR) <> Empty Then
            Call checkAndFormat(QCCurRow, RegRR, Regulatory_Admin, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, RegRR), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, RegRR), 2)
        End If
        
        If Cells(QCCurRow, BadDtCol) <> Empty Then
            Call checkAndFormat(QCCurRow, BadDtCol, Bad_Debt, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, BadDtCol), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, BadDtCol), 2)
        End If
        
        If Cells(QCCurRow, SDF) <> Empty Then
            Call checkAndFormat(QCCurRow, SDF, Electric_Service_Delivery, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, SDF), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, SDF), 2)
        End If
        
        If Cells(QCCurRow, SDF2) <> Empty Then
            Call checkAndFormat(QCCurRow, SDF2, Electric_Service_Delivery_2, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, SDF2), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, SDF2), 2)
        End If
            
        If Cells(QCCurRow, ElecMeter) <> Empty Then
            If Cells(1, ElecMeterCons) = Empty Then 'Flat
                Call checkAndFormat(QCCurRow, ElecMeter, Electric_Meter, CycleLength, ProLength)
            End If
            If Cells(1, ElecMeterCons) <> Empty Then 'Metered
                x = Round(checkMeteredCharge(ElecMeter_Rate, Cells(QCCurRow, ElecMeterCons)), 2)
                If Cells(QCCurRow, ElecMeter) = x Then
                    Cells(QCCurRow, ElecMeter).Select
                    turnGreen
                ElseIf Cells(QCCurRow, ElecMeter) <= (x + 0.01) And Cells(QCCurRow, ElecMeter) >= (x - 0.01) Then
                    Cells(QCCurRow, ElecMeter).Select
                    tintGreen
                ElseIf Cells(QCCurRow, ElecMeter) <= (x + 0.02) And Cells(QCCurRow, ElecMeter) >= (x - 0.02) Then
                    Cells(QCCurRow, ElecMeter).Select
                    warnOrange
                Else
                    Cells(QCCurRow, ElecMeter).Select
                    turnRed
                End If
            End If
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ElecMeter), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, ElecMeter), 2)
        End If
        
        If Cells(QCCurRow, ElecMeter2) <> Empty Then
            Call checkAndFormat(QCCurRow, ElecMeter2, Electric_Meter2, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ElecMeter2), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, ElecMeter2), 2)
        End If
        
        If Cells(QCCurRow, ElecMeter3) <> Empty Then
            Call checkAndFormat(QCCurRow, ElecMeter3, Electric_Meter3, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ElecMeter3), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, ElecMeter3), 2)
        End If
        
        If Cells(QCCurRow, ElecMeter4) <> Empty Then
            Call checkAndFormat(QCCurRow, ElecMeter4, Electric_Meter4, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ElecMeter4), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, ElecMeter4), 2)
        End If
        
        If Cells(QCCurRow, ElecMeter5) <> Empty Then
            Call checkAndFormat(QCCurRow, ElecMeter5, Electric_Meter5, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ElecMeter5), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, ElecMeter5), 2)
        End If
        
        If Cells(QCCurRow, COVIDRecCol) <> Empty Then
            Call checkAndFormat(QCCurRow, COVIDRecCol, COVID_Rec, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, COVIDRecCol), 2)
            WyseOERAccumulator = WyseOERAccumulator + Round(Cells(QCCurRow, COVIDRecCol), 2)
        End If
        
        If Cells(QCCurRow, WWA) <> Empty Then
            Call checkAndFormat(QCCurRow, WWA, Water_Service, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WWA), 2)
        End If
            
        If Cells(QCCurRow, WtrMtr) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr, Water_Meter, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr2) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr2, Water_Meter_2, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr2), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr3) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr3, Water_Meter_3, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr3), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr4) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr4, Water_Meter_4, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr4), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr5) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr5, Water_Meter_5, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr5), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr6) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr6, Water_Meter_6, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr6), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr7) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr7, Water_Meter_7, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr7), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr8) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr8, Water_Meter_8, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr8), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr9) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr9, Water_Meter_9, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr9), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr10) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr10, Water_Meter_10, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr10), 2)
        End If
        
        If Cells(QCCurRow, WtrMtr11) <> Empty Then
            Call checkAndFormat(QCCurRow, WtrMtr11, Water_Meter_11, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, WtrMtr11), 2)
        End If
        
        If Cells(QCCurRow, StormWaterSewerCol) <> Empty Then
            Call checkAndFormat(QCCurRow, StormWaterSewerCol, StrmWtrSewer, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, ThrmAdCol) <> Empty Then
            Call checkAndFormat(QCCurRow, ThrmAdCol, Thermal_Admin, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThrmAdCol), 2)
        End If
        
        If Cells(QCCurRow, ThrmMCol) <> Empty Then
            Call checkAndFormat(QCCurRow, ThrmMCol, Thermal_Meter, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThrmMCol), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter2_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter2_Col, Thermal_Meter_2, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter2_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter3_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter3_Col, Thermal_Meter_3, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter3_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter4_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter4_Col, Thermal_Meter_4, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter4_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter5_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter5_Col, Thermal_Meter_5, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter5_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter6_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter6_Col, Thermal_Meter_6, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter6_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter7_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter7_Col, Thermal_Meter_7, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter7_Col), 2)
        End If
        
        If Cells(QCCurRow, ThermalMeter8_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, ThermalMeter8_Col, Thermal_Meter_8, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Round(Cells(QCCurRow, ThermalMeter8_Col), 2)
        End If
        
        If Cells(QCCurRow, GasMeter_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, GasMeter_Col, gas_meter, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Cells(QCCurRow, GasMeter_Col)
        End If
        
        If Cells(QCCurRow, GasMeter2_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, GasMeter2_Col, gas_meter2, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Cells(QCCurRow, GasMeter2_Col)
        End If

        If Cells(QCCurRow, GasAdmin_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, GasAdmin_Col, gas_admin, CycleLength, ProLength)
            WyseHSTAccumulator = WyseHSTAccumulator + Cells(QCCurRow, GasAdmin_Col)
        End If
        
        ' RUBs Flat Fees
        If Cells(QCCurRow, ElecCust) <> Empty Then
            Call checkAndFormat(QCCurRow, ElecCust, elec_cust, CycleLength, ProLength)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, ElecCust), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, ElecCust), 2)
        End If
        
        If Cells(QCCurRow, Base) <> Empty Then
            Call checkAndFormat(QCCurRow, Base, Water_Base, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, SewerBase) <> Empty Then
            Call checkAndFormat(QCCurRow, SewerBase, Sewer_Base, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, FireSup) <> Empty Then
            Call checkAndFormat(QCCurRow, FireSup, Fire_Supply, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, CapCostRecCol) <> Empty Then
            Call checkAndFormat(QCCurRow, CapCostRecCol, Cap_Cst_Rec, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, StormWaterDrainCol) <> Empty Then
            Call checkAndFormat(QCCurRow, StormWaterDrainCol, StrmWtrDrain, CycleLength, ProLength)
        End If
        
        If Cells(QCCurRow, GasCust_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, GasCust_Col, gas_cust, CycleLength, ProLength)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, GasCust_Col)
        End If
        
        'Metered Fees

        ''Start Delivery Check
        If Cells(1, Delivcons) <> Empty Or Cells(1, MidPeakCons) <> Empty Then ''Metered
            If Cells(1, Deliv) <> Empty Then
                If Cells(QCCurRow, Energy) <> Empty And Cells(QCCurRow, MidPeak) = Empty And Cells(QCCurRow, OffPeak) = Empty And Cells(QCCurRow, OnPeak) = Empty Then
                    Call checkAndFormat(QCCurRow, Deliv, checkMeteredCharge(Delivery_Rate, Cells(QCCurRow, Delivcons)), 1, 1)
                    HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Deliv), 2)
                    OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Deliv), 2)
                End If
                
                If Cells(QCCurRow, Energy) = Empty And Cells(QCCurRow, MidPeak) <> Empty And Cells(QCCurRow, OffPeak) <> Empty And Cells(QCCurRow, OnPeak) <> Empty Then
                    Call checkAndFormat(QCCurRow, Deliv, checkMeteredCharge(Delivery_Rate, (Cells(QCCurRow, MidPeakCons) + Cells(QCCurRow, OffPeakCons) + Cells(QCCurRow, OnPeakCons)) * LAF), 1, 1)
                    HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Deliv), 2)
                    OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Deliv), 2)
                End If
            End If
        
            If answr2 = 6 Then
                '''Checks Regulatory Charge
                If Cells(QCCurRow, Energy) <> Empty And Cells(QCCurRow, MidPeak) = Empty And Cells(QCCurRow, OffPeak) = Empty And Cells(QCCurRow, OnPeak) = Empty Then
                    Call checkAndFormat(QCCurRow, Reg, checkMeteredCharge(RegChg_Rate, Cells(QCCurRow, RegCons)), 1, 1)
                    HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Reg)
                    OERAccumulator = OERAccumulator + Cells(QCCurRow, Reg)
                End If
                
                If Cells(QCCurRow, Energy) = Empty And Cells(QCCurRow, MidPeak) <> Empty And Cells(QCCurRow, OffPeak) <> Empty And Cells(QCCurRow, OnPeak) <> Empty Then
                    Call checkAndFormat(QCCurRow, Reg, checkMeteredCharge(RegChg_Rate, (Cells(QCCurRow, MidPeakCons) + Cells(QCCurRow, OffPeakCons) + Cells(QCCurRow, OnPeakCons)) * LAF), 1, 1)
                    HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Reg), 2)
                    OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Reg), 2)
                End If
                
                '''Checks LILO
                If Cells(1, LILOcons) <> Empty Or Cells(1, LILO) <> Empty Then  ''Metered
                
                    '''Tiered Billing LILO
                    If Cells(QCCurRow, Energy) <> Empty And Cells(QCCurRow, MidPeak) = Empty And Cells(QCCurRow, OffPeak) = Empty And Cells(QCCurRow, OnPeak) = Empty Then
                        Call checkAndFormat(QCCurRow, LILO, checkMeteredCharge(Energy_Rate, Cells(QCCurRow, LILOcons)), 1, 1)
                        HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, LILO), 2)
                        OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, LILO), 2)
                    End If
                
                    '''TOU Billing LILO
                    If Cells(QCCurRow, Energy) = Empty And Cells(QCCurRow, MidPeak) <> Empty And Cells(QCCurRow, OffPeak) <> Empty And Cells(QCCurRow, OnPeak) <> Empty Then
                        Call checkAndFormat(QCCurRow, LILO, (checkMeteredCharge(MidPeak_Rate, (Cells(QCCurRow, MidPeakCons) * (LAF - 1))) + checkMeteredCharge(OffPeak_Rate, (Cells(QCCurRow, OffPeakCons) * (LAF - 1))) + checkMeteredCharge(OnPeak_Rate, (Cells(QCCurRow, OnPeakCons) * (LAF - 1)))), 1, 1)
                        HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, LILO), 2)
                        OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, LILO), 2)
                    End If
                End If
            End If
        End If
        ''End elecdel cons check block
        
        '''Checks Energy Charges (Not yet working as Multi-Meter)
        If Cells(QCCurRow, Energy) <> Empty Then
            Call checkAndFormat(QCCurRow, Energy, checkMeteredCharge(Energy_Rate, Cells(QCCurRow, EnCons)), 1, 1)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Energy), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Energy), 2)
        End If
        
        If Cells(QCCurRow, Energy_2) <> Empty Then
            If E1E2_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, Energy_2, checkMeteredCharge(Energy2_Rate, Cells(QCCurRow, EnCons_2), Cells(QCCurRow, EnCons)), 1, 1)
            Else
                Call checkAndFormat(QCCurRow, Energy_2, checkMeteredCharge(Energy2_Rate, Cells(QCCurRow, EnCons_2)), 1, 1)
            End If
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Energy_2), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Energy_2), 2)
        End If
        
        If Cells(QCCurRow, Energy_3) <> Empty Then
            If E1E3_MM_Answer = 6 And E2E3_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, Energy_3, checkMeteredCharge(Energy3_Rate, Cells(QCCurRow, EnCons_3), Cells(QCCurRow, EnCons) + Cells(QCCurRow, EnCons_2)), 1, 1)
            ElseIf E1E3_MM_Answer = 6 And E2E3_MM_Answer = 7 Then
                Call checkAndFormat(QCCurRow, Energy_3, checkMeteredCharge(Energy3_Rate, Cells(QCCurRow, EnCons_3), Cells(QCCurRow, EnCons)), 1, 1)
            ElseIf E1E3_MM_Answer = 7 And E2E3_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, Energy_3, checkMeteredCharge(Energy3_Rate, Cells(QCCurRow, EnCons_3), Cells(QCCurRow, EnCons_2)), 1, 1)
            Else
                Call checkAndFormat(QCCurRow, Energy_3, checkMeteredCharge(Energy3_Rate, Cells(QCCurRow, EnCons_3)), 1, 1)
            End If
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, Energy_3), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, Energy_3), 2)
        End If
    
        '''Checks TOU Charges
        If Cells(QCCurRow, MidPeak) <> Empty Then  ''Metered
            Call checkAndFormat(QCCurRow, MidPeak, checkMeteredCharge(MidPeak_Rate, Cells(QCCurRow, MidPeakCons)), 1, 1)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, MidPeak), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, MidPeak), 2)
        End If
        
        If Cells(QCCurRow, OffPeak) <> Empty Then  ''Metered
            Call checkAndFormat(QCCurRow, OffPeak, checkMeteredCharge(OffPeak_Rate, Cells(QCCurRow, OffPeakCons)), 1, 1)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, OffPeak), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, OffPeak), 2)
        End If
        
        If Cells(QCCurRow, OnPeak) <> Empty Then  ''Metered
            Call checkAndFormat(QCCurRow, OnPeak, checkMeteredCharge(OnPeak_Rate, Cells(QCCurRow, OnPeakCons)), 1, 1)
            HSTAccumulator = HSTAccumulator + Round(Cells(QCCurRow, OnPeak), 2)
            OERAccumulator = OERAccumulator + Round(Cells(QCCurRow, OnPeak), 2)
        End If
        
        'Check Thermal Chg
        If Cells(QCCurRow, ThrmCol) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, ThrmCol, checkMeteredCharge(Thermal_Rate, Cells(QCCurRow, ThrmCnCol)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, ThrmCol)
        End If
        
        If Cells(QCCurRow, Cooling_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Cooling_Col, checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, CoolingCon_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Cooling_Col)
        End If
        
        If Cells(QCCurRow, Cooling2_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Cooling2_Col, checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling2Con_Col)), 1, 1)
            x = Round(checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling2Con_Col)), 2) ' New Code
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Cooling2_Col)
        End If
        
        If Cells(QCCurRow, Cooling3_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Cooling3_Col, checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling3Con_Col)), 1, 1)
            x = Round(checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling3Con_Col)), 2) ' New Code
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Cooling3_Col)
        End If
        
        If Cells(QCCurRow, Cooling4_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Cooling4_Col, checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling4Con_Col)), 1, 1)
            x = Round(checkMeteredCharge(Cooling_Rate, Cells(QCCurRow, Cooling4Con_Col)), 2) ' New Code
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Cooling4_Col)
        End If
        
        If Cells(QCCurRow, Heating_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Heating_Col, checkMeteredCharge(Heating_Rate, Cells(QCCurRow, HeatingCon_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Heating_Col)
        End If
        
        If Cells(QCCurRow, Heating2_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Heating2_Col, checkMeteredCharge(Heating_Rate, Cells(QCCurRow, Heating2Con_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Heating2_Col)
        End If
        
        If Cells(QCCurRow, Heating3_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Heating3_Col, checkMeteredCharge(Heating_Rate, Cells(QCCurRow, Heating3Con_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Heating3_Col)
        End If
        
        If Cells(QCCurRow, Heating4_Col) <> Empty Then ''Metered
            Call checkAndFormat(QCCurRow, Heating4_Col, checkMeteredCharge(Heating_Rate, Cells(QCCurRow, Heating4Con_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Heating4_Col)
        End If

        '''Start Water Block
        If Cells(QCCurRow, wtR) <> Empty Then
            Call checkAndFormat(QCCurRow, wtR, checkMeteredCharge(Water_Rate, Cells(QCCurRow, wtRcons)), 1, 1)
        End If
        
        If Cells(QCCurRow, wtR2) <> Empty Then
            If W1W2_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, wtR2, checkMeteredCharge(Water2_Rate, Cells(QCCurRow, wtR2cons), Cells(QCCurRow, wtRcons)), 1, 1)
            Else
                Call checkAndFormat(QCCurRow, wtR2, checkMeteredCharge(Water2_Rate, Cells(QCCurRow, wtR2cons)), 1, 1)
            End If
        End If
        
        If Cells(QCCurRow, wtR3) <> Empty Then
            If W1W3_MM_Answer = 6 And W2W3_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, wtR3, checkMeteredCharge(Water3_Rate, Cells(QCCurRow, wtR3cons), Cells(QCCurRow, wtRcons) + Cells(QCCurRow, wtR2cons)), 1, 1)
            ElseIf W1W3_MM_Answer = 6 And W2W3_MM_Answer = 7 Then
                Call checkAndFormat(QCCurRow, wtR3, checkMeteredCharge(Water3_Rate, Cells(QCCurRow, wtR3cons), Cells(QCCurRow, wtRcons)), 1, 1)
            ElseIf W1W3_MM_Answer = 7 And W2W3_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, wtR3, checkMeteredCharge(Water3_Rate, Cells(QCCurRow, wtR3cons), Cells(QCCurRow, wtR2cons)), 1, 1)
            Else
                Call checkAndFormat(QCCurRow, wtR3, checkMeteredCharge(Water3_Rate, Cells(QCCurRow, wtR3cons)), 1, 1)
            End If
        End If
            
        If Cells(QCCurRow, coldwtR) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, coldwtR, checkMeteredCharge(clDWater_Rate, Cells(QCCurRow, coldwtRcons)), 1, 1)
        End If
        
        If Cells(QCCurRow, coldwtR2) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, coldwtR2, checkMeteredCharge(clDWater2_Rate, Cells(QCCurRow, coldwtR2cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, coldwtR3) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, coldwtR3, checkMeteredCharge(clDWater3_Rate, Cells(QCCurRow, coldwtR3cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, coldwtR4) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, coldwtR4, checkMeteredCharge(clDWater4_Rate, Cells(QCCurRow, coldwtR4cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, coldwtR5) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, coldwtR5, checkMeteredCharge(clDWater5_Rate, Cells(QCCurRow, coldwtR5cons)), 1, 1)
        End If

        If Cells(QCCurRow, hotwtR) <> Empty Then ' new Code
            If CW1HW1_MM_Answer = 6 Then
                Call checkAndFormat(QCCurRow, hotwtR, checkMeteredCharge(hTWater_Rate, Cells(QCCurRow, hotwtRcons), Cells(QCCurRow, coldwtRcons)), 1, 1)
            Else
                Call checkAndFormat(QCCurRow, hotwtR, checkMeteredCharge(hTWater_Rate, Cells(QCCurRow, hotwtRcons)), 1, 1)
            End If
        End If
        
        If Cells(QCCurRow, hotwtR2) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, hotwtR2, checkMeteredCharge(hTWater2_Rate, Cells(QCCurRow, hotwtR2cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, hotwtR3) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, hotwtR3, checkMeteredCharge(hTWater3_Rate, Cells(QCCurRow, hotwtR3cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, hotwtR4) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, hotwtR4, checkMeteredCharge(hTWater4_Rate, Cells(QCCurRow, hotwtR4cons)), 1, 1)
        End If
        
        If Cells(QCCurRow, hotwtR5) <> Empty Then ' new Code
            Call checkAndFormat(QCCurRow, hotwtR5, checkMeteredCharge(hTWater5_Rate, Cells(QCCurRow, hotwtR5cons)), 1, 1)
        End If
            
        '''End Water Block
    
            ''Sewer
        If Cells(QCCurRow, Sewer) <> Empty Then
            Call checkAndFormat(QCCurRow, Sewer, checkMeteredCharge(Sewer_Rate, Cells(QCCurRow, SewerCons)), 1, 1)
        End If
        
        If Cells(QCCurRow, Sewer2) <> Empty Then
            Call checkAndFormat(QCCurRow, Sewer2, checkMeteredCharge(Sewer2_Rate, Cells(QCCurRow, Sewer2Cons)), 1, 1)
        End If
        
        'Gas
        If Cells(QCCurRow, Gas_Col) <> Empty Then
            Call checkAndFormat(QCCurRow, Gas_Col, checkMeteredCharge(Gas_Rate, Cells(QCCurRow, GasCon_Col)), 1, 1)
            HSTAccumulator = HSTAccumulator + Cells(QCCurRow, Gas_Col)
        End If
        
        ''Deposit
        If Cells(QCCurRow, UTILDEP) <> Empty Then
            If Cells(QCCurRow, UTILDEP) = Cells(QCCurRow, DepAmt) Then
                Cells(QCCurRow, UTILDEP).Select
                turnGreen
            End If
        End If
        
        ''Paper Bill Fee
        
        If Cells(1, EbillAddie_Col) <> "" Then
            If Cells(QCCurRow, EbillAddie_Col) <> "" Then
                If Cells(QCCurRow, PaperBillTax_Col) = "" Then
                    Union(Cells(QCCurRow, EbillAddie_Col), Cells(QCCurRow, PaperBillTax_Col)).Select
                    turnGreen
                Else
                    Union(Cells(QCCurRow, EbillAddie_Col), Cells(QCCurRow, PaperBillTax_Col)).Select
                    turnRed
                End If
            ElseIf Cells(QCCurRow, EbillAddie_Col) = "" Then
                If Cells(QCCurRow, PaperBillTax_Col) = 2 Then
                    Union(Cells(QCCurRow, EbillAddie_Col), Cells(QCCurRow, PaperBillTax_Col)).Select
                    turnGreen
                Else
                    Union(Cells(QCCurRow, EbillAddie_Col), Cells(QCCurRow, PaperBillTax_Col)).Select
                    turnRed
                End If
            End If
        WyseHSTAccumulator = WyseHSTAccumulator + Cells(QCCurRow, PaperBillTax_Col)
        End If
        
        'Check PAs
        If Cells(QCCurRow, HasPyArr_Col) = "No" Then
            If Cells(1, PastDueArr_Col) = "" Or Cells(QCCurRow, PastDueArr_Col) = "" Then
                Cells(QCCurRow, HasPyArr_Col).Select
                turnGreen
                If Cells(1, PastDueArr_Col) <> "" Then
                    Cells(QCCurRow, PastDueArr_Col).Select
                    turnGreen
                End If
            Else
                Union(Cells(QCCurRow, HasPyArr_Col), Cells(QCCurRow, PastDueArr_Col)).Select
                turnRed
            End If
        ElseIf Cells(QCCurRow, HasPyArr_Col) = "Yes" Then
            If Cells(QCCurRow, PAMonthly_Col) = Cells(QCCurRow, PastDueArr_Col) Then
                Union(Cells(QCCurRow, PAMonthly_Col), Cells(QCCurRow, PastDueArr_Col)).Select
                turnGreen
            Else
                Union(Cells(QCCurRow, PAMonthly_Col), Cells(QCCurRow, PastDueArr_Col)).Select
                turnRed
            End If
            
            If Cells(QCCurRow, PrevBal_Col) <= 0 Then
                Cells(QCCurRow, HasPyArr_Col).Select
                turnGreen
            Else
                Cells(QCCurRow, HasPyArr_Col).Select
                turnRed
            End If
        End If
        
        'Only runs if Yes to Ontario Prop
        If answr2 = 6 Then
        
            ''Turns Colors OESP tax credits if they have been imported
            If Cells(1, OESPHSTCR) <> Empty Then
                If Cells(QCCurRow, OESPHST) = -(Cells(QCCurRow, OESPHSTCR) + Cells(QCCurRow, HST) + Cells(QCCurRow, WyseHST) + Cells(QCCurRow, SUPHST)) Then
                    Cells(QCCurRow, OESPHST).Select
                    turnGreen
                End If
            ElseIf Cells(1, OESPHSTCR) = Empty And (Cells(QCCurRow, OESPHST) + Cells(QCCurRow, HST) + Cells(QCCurRow, WyseHST) + Cells(QCCurRow, SUPHST)) < 0 Then
                Cells(QCCurRow, XX).Value = -(Cells(QCCurRow, OESPHST) + Cells(QCCurRow, HST) + Cells(QCCurRow, WyseHST) + Cells(QCCurRow, SUPHST))
                Cells(QCCurRow, XX).Select
                turnRed
            End If
            
            '''Verifies HST
            If Cells(1, HST) <> Empty Then
                check22 = (HSTAccumulator * HST_Rate)
                check22 = Round(check22, 2)
                If Cells(QCCurRow, HST) = check22 Then
                    Cells(QCCurRow, HST).Select
                    turnGreen
                ElseIf Cells(QCCurRow, HST) <= (check22 + 0.01) And Cells(QCCurRow, HST) >= (check22 - 0.01) Then
                    Cells(QCCurRow, HST).Select
                    tintGreen
                ElseIf Cells(QCCurRow, HST) <= (check22 + 0.02) And Cells(QCCurRow, HST) >= (check22 - 0.02) Then
                    Cells(QCCurRow, HST).Select
                   warnOrange
                Else
                    Cells(QCCurRow, HST).Select
                    turnRed
                End If
            End If
        
        '''Verifies Wyse HST
            If Cells(1, WyseHST) <> Empty Then
                check23 = (WyseHSTAccumulator * HST_Rate)
                check23 = Round(check23, 2)
                If Cells(QCCurRow, WyseHST) = check23 Then
                    Cells(QCCurRow, WyseHST).Select
                    turnGreen
                ElseIf Cells(QCCurRow, WyseHST) <= (check23 + 0.01) And Cells(QCCurRow, WyseHST) >= (check23 - 0.01) Then
                    Cells(QCCurRow, WyseHST).Select
                    tintGreen
                ElseIf Cells(QCCurRow, WyseHST) <= (check23 + 0.02) And Cells(QCCurRow, WyseHST) >= (check23 - 0.02) Then
                    Cells(QCCurRow, WyseHST).Select
                   warnOrange
                Else
                    Cells(QCCurRow, WyseHST).Select
                    turnRed
                End If
            End If
        
        '''Verifies Wyse Rebate (WyseOER)
            If Cells(1, wyseRbate) <> Empty Then
                check24 = (WyseOERAccumulator * OER_Rate)
                check24 = Round(check24, 2)
                If Cells(QCCurRow, wyseRbate) = check24 Then
                    Cells(QCCurRow, wyseRbate).Select
                    turnGreen
                ElseIf Cells(QCCurRow, wyseRbate) <= (check24 + 0.01) And Cells(QCCurRow, wyseRbate) >= (check24 - 0.01) Then
                    Cells(QCCurRow, wyseRbate).Select
                    tintGreen
                ElseIf Cells(QCCurRow, wyseRbate) <= (check24 + 0.02) And Cells(QCCurRow, wyseRbate) >= (check24 - 0.02) Then
                    Cells(QCCurRow, wyseRbate).Select
                    warnOrange
                Else
                    Cells(QCCurRow, wyseRbate).Select
                    turnRed
                End If
            End If
        
        '''Verifies Regular Rebate (OER)
            If Cells(1, RegRbate) <> Empty Then
                check25 = OERAccumulator * OER_Rate
                check25 = Round(check25, 2)
                If Cells(QCCurRow, RegRbate) = check25 Then
                    Cells(QCCurRow, RegRbate).Select
                    turnGreen
                ElseIf Cells(QCCurRow, RegRbate) <= (check25 + 0.01) And Cells(QCCurRow, RegRbate) >= (check25 - 0.01) Then
                    Cells(QCCurRow, RegRbate).Select
                    tintGreen
                ElseIf Cells(QCCurRow, RegRbate) <= (check25 + 0.02) And Cells(QCCurRow, RegRbate) >= (check25 - 0.02) Then
                    Cells(QCCurRow, RegRbate).Select
                    warnOrange
                Else
                    Cells(QCCurRow, RegRbate).Select
                    turnRed
                End If
            End If
    
            '''Verifies HST - OESP
            check27 = (Cells(QCCurRow, OESP).Value) * HST_Rate
            check27 = Round(check27, 2)
            
            If Cells(1, OESPHSTCR) = Empty Then
                If Cells(QCCurRow, OESPHST) = check27 Then
                    Cells(QCCurRow, OESPHST).Select
                    turnGreen
                ElseIf Cells(QCCurRow, OESPHST) <= (check27 + 0.01) And Cells(QCCurRow, OESPHST) >= (check27 - 0.01) Then
                    Cells(QCCurRow, OESPHST).Select
                    tintGreen
                ElseIf Cells(QCCurRow, OESPHST) <= (check27 + 0.02) And Cells(QCCurRow, OESPHST) >= (check27 - 0.02) Then
                    Cells(QCCurRow, OESPHST).Select
                    warnOrange
                End If
            End If
        
            If Cells(1, OESPHSTCR) <> Empty Then
                If Cells(QCCurRow, OESPHSTCR) = check27 Then
                    Cells(QCCurRow, OESPHSTCR).Select
                    turnGreen
                ElseIf Cells(QCCurRow, OESPHSTCR) <= (check27 + 0.01) And Cells(QCCurRow, OESPHSTCR) >= (check27 - 0.01) Then
                    Cells(QCCurRow, OESPHSTCR).Select
                    tintGreen
                ElseIf Cells(QCCurRow, OESPHSTCR) <= (check27 + 0.02) And Cells(QCCurRow, OESPHSTCR) >= (check27 - 0.02) Then
                    Cells(QCCurRow, OESPHSTCR).Select
                    warnOrange
                End If
            End If
        End If
        
    QCCurRow = QCCurRow + 1
    Loop


Call SS_RangeCheck_Loop(PropSS, PropSSshort, PropQC, XY)
        
End If

''Flag Names with EDE Non-Billable Tags
Check4NonBill

answr3 = MsgBox("Would you like to open the Spreadsheet?", vbYesNo)

If answr3 = 6 Then
    Workbooks.Open (PropSS)
End If

Exit Sub



End Sub


Function checkAndFormat(rezRow, chrgCol, chrgAmt, fulLength, shortLength)
    x = Round(((chrgAmt / fulLength) * shortLength), 2)
        If Cells(rezRow, chrgCol) = x Then
            Cells(rezRow, chrgCol).Select
            turnGreen
        ElseIf Cells(rezRow, chrgCol) >= (x - 0.01) And Cells(rezRow, chrgCol) <= (x + 0.01) Then
            Cells(rezRow, chrgCol).Select
            tintGreen
        ElseIf Cells(rezRow, chrgCol) >= (x - 0.02) And Cells(rezRow, chrgCol) <= (x + 0.02) Then
            Cells(rezRow, chrgCol).Select
            warnOrange
        Else
            Cells(rezRow, chrgCol).Select
            turnRed
        End If
End Function

Function Find_info(Rate, n, txt, txt2, Optional errorcol)
    '''for small #'s with many decimal places
    On Error GoTo Find_Info_Error
    Dim wbName As String, wb As Workbook, ws As Worksheet
    Dim eXtra As Double
    Dim Q As Integer
    Dim List_Names() As Variant
    Dim currow As Integer
    Dim Name_Count As Integer

    Application.ScreenUpdating = False
    Workbooks.Open (txt)

    eXtra = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Cells(1, n)

    Workbooks(txt2).Close
    Find_info = eXtra
    Exit Function
    
Find_Info_Error:

Workbooks.Open ("U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\NamedRangeList.xlsx")

currow = 1
Do Until Cells(currow, 1) = ""
    currow = currow + 1
Loop

ReDim List_Names(1 To currow, 2) As Variant
Name_Count = UBound(List_Names) - LBound(List_Names) + 1
For x = 1 To Name_Count
    List_Names(x, 1) = Cells(x, 1)
    List_Names(x, 2) = Cells(x, 2)
Next x
    
ActiveWorkbook.Close
ActiveWorkbook.Close

For Rowe = 1 To Name_Count
    If Cells(1, errorcol) = List_Names(Rowe, 2) Then
        Output1 = List_Names(Rowe, 2)
        Output2 = List_Names(Rowe, 1)
        Workbooks.Open (txt)
        Report = MsgBox("Cannot check the charge " & Output1 & ".  Please add the named range " & Output2 & " to your property spreadsheet and retry.", vbOKOnly)
        End
    End If
Next Rowe

End Function

Function FindRUBs_info(Rate, txt, txt2, Optional errorcol)
    '''for small #'s with many decimal places
    On Error GoTo FindRUBS_Info_Error
    Dim wbName As String, wb As Workbook, ws As Worksheet
    Dim eXtra As Double
    Dim Q As Integer
    
    Application.ScreenUpdating = False
    Workbooks.Open (txt)

    n = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Count
        
    eXtra = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Cells(1, n)

    Workbooks(txt2).Close
    FindRUBs_info = eXtra
    Exit Function
    
FindRUBS_Info_Error:

Workbooks.Open ("U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\NamedRangeList.xlsx")

currow = 1
Do Until Cells(currow, 1) = ""
    currow = currow + 1
Loop

ReDim List_Names(1 To currow, 2) As Variant
Name_Count = UBound(List_Names) - LBound(List_Names) + 1
For x = 1 To Name_Count
    List_Names(x, 1) = Cells(x, 1)
    List_Names(x, 2) = Cells(x, 2)
Next x
    
ActiveWorkbook.Close
ActiveWorkbook.Close

For Rowe = 1 To Name_Count
    If Cells(1, errorcol) = List_Names(Rowe, 2) Then
        Output1 = List_Names(Rowe, 2)
        Output2 = List_Names(Rowe, 1)
        Workbooks.Open (txt)
        Report = MsgBox("Cannot check the charge " & Output1 & ".  Please add the named range " & Output2 & " to your property spreadsheet and retry.", vbOKOnly)
        End
    End If
Next Rowe

End Function

Function Find_rate(Rate, txt, txt2, Optional errorcol)
    Dim IntX As Integer, IntY As Integer
    Dim RtTiers As Long
        
    On Error GoTo Find_rate_Error
    Application.ScreenUpdating = False
    Workbooks.Open (txt)
    
    RtTiers = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Rows.Count
    
    ReDim eXtra(1 To RtTiers, 3) As Variant
        
    If RtTiers = 1 Then 'single-tiered metered charge
        eXtra(1, 1) = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Cells(1, 1)
        eXtra(1, 2) = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Cells(1, 2)
        eXtra(1, 3) = "Single-Tier"
    Else 'multi-tiered metered charge
        For IntX = 1 To RtTiers
            For IntY = 1 To 3
                eXtra(IntX, IntY) = Workbooks(txt2).Worksheets("Bills - TCC").Range(Rate).Cells(IntX, IntY)
            Next IntY
        Next IntX
    End If
    
    Workbooks(txt2).Close
    Find_rate = eXtra
    Exit Function
    
Find_rate_Error:

Workbooks.Open ("U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\NamedRangeList.xlsx")

currow = 1
Do Until Cells(currow, 1) = ""
    currow = currow + 1
Loop

ReDim List_Names(1 To currow, 2) As Variant
Name_Count = UBound(List_Names) - LBound(List_Names) + 1
For x = 1 To Name_Count
    List_Names(x, 1) = Cells(x, 1)
    List_Names(x, 2) = Cells(x, 2)
Next x
    
ActiveWorkbook.Close
ActiveWorkbook.Close

For Rowe = 1 To Name_Count
    If Cells(1, errorcol) = List_Names(Rowe, 2) Then
        Output1 = List_Names(Rowe, 2)
        Output2 = List_Names(Rowe, 1)
        Workbooks.Open (txt)
        Report = MsgBox("Cannot check the charge " & Output1 & ".  Please add the named range " & Output2 & " to your property spreadsheet and retry.", vbOKOnly)
        End
    End If
Next Rowe

End Function

Function checkMeteredCharge(Rate, utilCons, Optional PrevUteCons As Double = 0)
    Dim UtilX As Integer
    Dim UtilY As Integer
    Dim RtTiers As Integer
    
    RtTiers = UBound(Rate) - LBound(Rate) + 1
    
    If RtTiers = 1 Then 'Single-Tiered metered charge
        eXtra = utilCons * Rate(1, 2)
    ElseIf PrevUteCons <> 0 Then
        For UtilX = 1 To RtTiers
            If UtilX = 1 Then
                If PrevUteCons > Rate(UtilX, 3) Then
                    curTierCons = 0
                ElseIf (Rate(UtilX, 3) - PrevUteCons) > utilCons Then
                    curTierCons = utilCons
                    eXtra = eXtra + (Rate(UtilX, 2) * curTierCons)
                Else
                    curTierCons = Rate(UtilX, 3) - PrevUteCons
                    eXtra = eXtra + (Rate(UtilX, 2) * curTierCons)
                End If
            ElseIf UtilX > 1 And UtilX < RtTiers Then
                If curTierCons > Rate(UtilX, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * (Rate(UtilX, 3) - Rate(UtilX - 1, 3)))
                    curTierCons = curTierCons - (Rate(UtilX, 3) - Rate(UtilX - 1, 3))
                ElseIf curTierCons < Rate(UtilX, 3) And curTierCons >= Rate(UtilX - 1, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * (curTierCons - Rate(UtilX - 1, 3)))
                End If
            ElseIf UtilX = RtTiers Then
                curTierCons = utilCons - curTierCons
                If curTierCons > 0 Then
                    eXtra = eXtra + (Rate(UtilX, 2) * curTierCons)
                End If
            End If
        Next UtilX
    Else
        For UtilX = 1 To RtTiers 'Multi-Tiered metered charge
            If UtilX = 1 Then
                If utilCons > Rate(UtilX, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * Rate(UtilX, 3))
                Else
                    eXtra = eXtra + (Rate(UtilX, 2) * utilCons)
                End If
            ElseIf UtilX > 1 And UtilX < RtTiers Then
                If utilCons > Rate(UtilX, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * (Rate(UtilX, 3) - Rate(UtilX - 1, 3)))
                ElseIf utilCons < Rate(UtilX, 3) And utilCons >= Rate(UtilX - 1, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * (utilCons - Rate(UtilX - 1, 3)))
                End If
            ElseIf UtilX = RtTiers Then
                If utilCons > Rate(UtilX - 1, 3) Then
                    eXtra = eXtra + (Rate(UtilX, 2) * (utilCons - Rate(UtilX - 1, 3)))
                End If
            End If
        Next UtilX
    End If
    
    checkMeteredCharge = eXtra

End Function

Function findPropSpreadsheet(shortName)
    Dim addy As String
    Dim Propco As String
    Dim Year As String
    Dim BM As String
    Dim currow1 As Integer
    Dim BMFind As String
    Dim length As Integer
    Dim BMfirst As Integer
    Dim PCFIND As Integer
    Dim PCFINDL As Integer
    Dim FindHMY As String
    
    BMFind = InStr(Cells(1, 1), "- ")
    length = Len(Cells(1, 1))
    BMFind = Mid(Cells(1, 1), BMFind + 2, length)
    BMFind = CDate(BMFind)
    Cells(3, 1) = BMFind
    bmlast = Right(BMFind, 2)
    Cells(3, 1) = Cells(3, 1) + 8
    BMFind = Cells(3, 1)
    BMFind = CDate(BMFind)
    BMfirst = Month(BMFind)

    PCFIND = InStr(Cells(2, 1), "(")
    PCFINDL = Len(Cells(2, 1))
    If PCFINDL - PCFIND = 6 Then
        Propco = Right(Cells(2, 1), 6)
        Propco = Left(Propco, 5)
    End If
    If PCFINDL - PCFIND = 5 Then
        Propco = Right(Cells(2, 1), 5)
        Propco = Left(Propco, 4)
    End If

    If Propco = "rt46" Or Propco = "lw08" Or Propco = "wc38" Then
        BMfirst = Month(BMFind) - 1
    End If

    Cells(3, 1) = BMfirst & bmlast
    Cells(3, 1).NumberFormat = "0000"

    addy = ActiveWorkbook.Worksheets("Summary of Utilities Billed - Q").Cells(2, 1).Value
    Year = Right(Cells(1, 1), 4)
    BM = Cells(3, 1)
    If Len(BM) < 4 Then
        BM = "0" & BM
    End If
    currow1 = 1

    Application.ScreenUpdating = False
    Workbooks.Open ("U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\HMY.xlsx")

    currow1 = 1
    Do Until Cells(currow1, 1) = Propco
        currow1 = currow1 + 1
    Loop
    FindHMY = CStr(Cells(currow1, 2))

    findPropSpreadsheet = "\\clientfiles\Properties\" & FindHMY & "\Billing\" & Year & "\" & BM & "\" & Propco & " Bill Spreadsheet " & BM & ".xlsx"
    shortName = Propco & " Bill Spreadsheet " & BM & ".xlsx"

    Workbooks("HMY.xlsx").Close

End Function

Function SS_RangeCheck_Loop(SStxt, SStxtShort, QCtxt, QCSumEnd)
    Dim nm As Name
    Dim List_Names() As Variant
    Dim currow As Integer
    Dim Name_Count As Integer
    Dim NameX As Integer
                
    Application.ScreenUpdating = False
    Workbooks.Open ("U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\NamedRangeList.xlsx")
    
    currow = 1
    Do Until Cells(currow, 1) = ""
        currow = currow + 1
    Loop
    
    ReDim List_Names(1 To currow, 2) As Variant
    Name_Count = UBound(List_Names) - LBound(List_Names) + 1
    
    For x = 1 To Name_Count
        List_Names(x, 1) = Cells(x, 1)
        List_Names(x, 2) = Cells(x, 2)
    Next x
        
    ActiveWorkbook.Close
    
    Workbooks.Open (SStxt)
    
    For Each nm In ActiveWorkbook.Names
        For NameX = 1 To Name_Count
            If Not nm.Name Like "_xlfn*" Then
                If nm.Name <> "LAF" Then
                    If nm.Name = List_Names(NameX, 1) Then
                        nameCheck = List_Names(NameX, 2)
                        isPresent = False
                        Workbooks(QCtxt).Activate
                        For y = 1 To QCSumEnd
                            If Cells(1, y) = nameCheck Then
                                isPresent = True
                            End If
                        Next y
                        Workbooks(SStxtShort).Activate
                        If isPresent = False Then
                            If Output = "" Then
                                Output = List_Names(NameX, 2)
                            Else
                                Output = Output & ", " & List_Names(NameX, 2)
                            End If
                        End If
                    End If
                End If
            End If
        Next NameX
    Next nm
    
    ActiveWorkbook.Close
        
    If Output <> "" Then
        Report = MsgBox("Spreadsheet has Named Ranges for the following charges that are not billed:" & Chr(13) & Chr(13) & Output & Chr(13) & Chr(13) & "Please verify that these charges should not be billed.", vbOKOnly)
    End If
    
End Function

Function Array_Compare(Array1, Array2)
    Dim Array1Size
    Dim Array2Size
    Dim Answer As Integer
    Dim RatesMatch As Boolean
    
    Array1Size = UBound(Array1, 1) * UBound(Array1, 2)
    Array2Size = UBound(Array2, 1) * UBound(Array2, 2)
    
    If Array1Size = Array2Size Then
        For x = 1 To UBound(Array1, 1)
            For y = 1 To UBound(Array1, 2)
                If Array1(x, y) <> Array2(x, y) Then
                    RatesMatch = False
                End If
            Next y
        Next x
        RatesMatch = True
    Else
        RatesMatch = False
    End If
    
    Array_Compare = RatesMatch
End Function

Function closeSS(txt)
    Dim wb As Workbook

    For Each wb In Workbooks
        If wb.Name = txt Then
            wb.Close
        End If
    Next
End Function

Function Check4NonBill()
    Dim checkRow As Integer
    Dim IntX As Integer
    Dim NumTags As Integer
    Dim Addie As Variant
    Dim ShortAddie As Variant
    
    Application.ScreenUpdating = False
    Addie = "U:\Conservice\Client files\Wyse Meter Solutions\Z-Wyse Billing Info\Auto QC\EDE NonBillable Tags.xlsx"
    ShortAddie = "EDE NonBillable Tags.xlsx"
    checkRow = 1
    Workbooks.Open (Addie)
    
    NumTags = Workbooks(ShortAddie).Worksheets("Sheet1").Range("Tags").Rows.Count
    
    ReDim eXtra(1 To NumTags) As Variant
    
    For IntX = 1 To NumTags
        eXtra(IntX) = Workbooks(ShortAddie).Worksheets("Sheet1").Range("Tags").Cells(IntX)
    Next IntX
    
    Workbooks(ShortAddie).Close
    
    Do Until Cells(checkRow, 1) = Empty
        For IntX = 1 To NumTags
            If InStr(Cells(checkRow, 2), eXtra(IntX)) <> 0 Then
                Cells(checkRow, 2).Select
                turnRed
            End If
        Next IntX
        checkRow = checkRow + 1
    Loop
End Function
 
 

Function turnGreen()
    With selection.Interior
        .Color = RGB(146, 208, 80)
        .TintAndShade = 0
    End With
End Function

Function tintGreen()
    With selection.Interior
        .Color = RGB(198, 224, 180)
        .TintAndShade = 0
    End With
End Function

Function warnOrange()
    With selection.Interior
        .Color = RGB(242, 157, 100)
        .TintAndShade = 0
    End With
End Function

Function turnRed()
    With selection.Interior
        .Color = RGB(255, 143, 156)
        .TintAndShade = 0
    End With
End Function

Function tintNMI()
    With selection.Interior
        .Color = RGB(209, 193, 221)
        .TintAndShade = 0
    End With
End Function





