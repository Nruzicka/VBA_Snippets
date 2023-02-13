Public Function PERCENTCHANGE(newval As Integer, oldval As Integer)
    PERCENTCHANGE = (newval - oldval) / oldval
End Function

Public Function ExpSmooth(hist As Range, alpha As Single)
    Dim QTY As Variant
    QTY = arrayManip.rangeToArray(hist)
    
    Dim initMean As Single
    initMean = QTY(0)
    
    'building Base
    Dim base() As Single
    ReDim base(UBound(QTY) + 1)
    base(0) = initMean
    
    'loop and calculate base
    For f = 1 To UBound(base)
        base(f) = (alpha * QTY(f - 1)) + ((1 - alpha) * (base(f - 1)))
    Next f

    ExpSmooth = base(UBound(base))
End Function

Function Holt(hist As Range, alpha As Single, beta As Single)
    Dim QTY As Variant
    QTY = arrayManip.rangeToArray(hist)

'store avg of QTY array into var "initial mean"  (NOTE: in other methods, the first data point is used)
    Dim initMean As Single
    initMean = WorksheetFunction.Average(QTY)
     
'create slope array from QTY array
    Dim slope() As Variant
    ReDim slope(UBound(QTY) - 1)
    
    For j = 0 To UBound(QTY) - 1
        slope(j) = QTY(j + 1) - QTY(j)
    Next j
      
'store avg of slope array in var "Initial tend"
    Dim inittrend As Single
    inittrend = WorksheetFunction.Average(slope)
     
'create base array and place initmean as first unit. Make array len of qty+1
    Dim base() As Single
    ReDim base(UBound(QTY) + 1)
    base(0) = initMean
    
'create trend array and place inittrend as first unit. Make array len of qty+1
    Dim trend() As Single
    ReDim trend(UBound(QTY) + 1)
    trend(0) = inittrend
    
'Create fcst array and place sum of initmean and trend. Make array len of qty+1
    Dim fcst() As Single
    ReDim fcst(UBound(QTY) + 1)
    fcst(0) = initMean + inittrend
     
'For len QTY+1, execute Holt algorithm
    For f = 1 To UBound(QTY) + 1
        base(f) = (alpha * QTY(f - 1)) + ((1 - alpha) * (base(f - 1) + trend(f - 1)))
        trend(f) = (beta * (base(f) - base(f - 1))) + ((1 - beta) * trend(f - 1))
        fcst(f) = base(f) + trend(f)
    Next f
    
    Holt = fcst(UBound(fcst))
End Function
