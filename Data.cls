VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' min
Private pmin As Double
' max
Private pmax As Double

' mean
Private xbar As Double

' variance
Private var As Double

' stddev
Private standardDev As Double

' average deveation
Private averageDev As Double

' the data we're going to calc stats for
Private data() As Double
' current write position in data()
Private pos As Long

Private Type pair
    val As Double
    one As Boolean
End Type

Public Property Get Count() As Long
    Count = pos
End Property

Public Property Get Item(idx As Long) As Double
Attribute Item.VB_UserMemId = 0
'
' returns data(idx)
'
    Item = data(idx)
End Property

Public Property Let Item(idx As Long, val As Double)
    data(idx) = val
End Property

Public Property Get Min() As Double
    Min = pmin
End Property

Public Property Get Max() As Double
    Max = pmax
End Property

Public Property Get StandardDeviation() As Double
    StandardDeviation = standardDev
End Property

Public Property Get Mean() As Double
    Mean = xbar
End Property

Public Property Get Variance() As Double
    Variance = var
End Property

Public Property Get AverageDeviation() As Double
    AverageDeviation = averageDev
End Property

Public Sub Clear()
'
' reset all values
'
    ReDim data(10)
    pos = 0
    xbar = 0
    var = 0
    pmin = 0
    pmax = 0
    standardDev = 0
    averageDev = 0
End Sub

Public Sub AppendData(val As Double)
'
' append a datum to data()
'
    data(pos) = val
    If pos = UBound(data) Then
        ReDim Preserve data(UBound(data) + 15)
    End If
    pos = pos + 1
End Sub

Public Sub CalcStats()
'
' calcualte the min, max, stddev, avgdev, variance, mean
'
    Dim x As Long
    
    pmin = 99999999999#
    pmax = -99999999999#
    xbar = 0#
    
    ' find the mean, min, max
    For x = 0 To pos - 1
        xbar = xbar + data(x)
        If data(x) > pmax Then
            pmax = data(x)
        End If
        If data(x) < pmin Then
            pmin = data(x)
        End If
    Next
    
    xbar = xbar / pos
    
    var = 0#
    averageDev = 0#
    
    ' calc the variance
    For x = 0 To pos - 1
        var = var + ((data(x) - xbar) * (data(x) - xbar))
        'var = var + data(x) * data(x)
        averageDev = averageDev + Abs(data(x) - xbar)
    Next
    'If pos = 1 Then
    '    var = 0
    'Else
        'var = var / (pos - 1)
        var = var / pos
    'End If
    
    'Debug.Assert (var - (xbar * xbar / pos)) / (pos - 1) >= 0#
    'var = (var - (xbar * xbar / pos)) / (pos - 1)
            
    standardDev = Sqr(var)
    averageDev = averageDev / pos
End Sub

Public Function Skew() As Double
    Dim x As Long
    Dim sum As Double
    
    sum = 0#
    For x = 0 To pos - 1
        sum = sum + ((data(x) - xbar) / standardDev) * ((data(x) - xbar) / standardDev) * ((data(x) - xbar) / standardDev)
    Next
    Skew = sum / pos
    
End Function

Public Function Kurtosis() As Double

    Dim x As Long
    Dim sum As Double
    
    sum = 0#
    For x = 0 To pos - 1
        sum = sum + ((data(x) - xbar) / standardDev) * ((data(x) - xbar) / standardDev) * ((data(x) - xbar) / standardDev) * ((data(x) - xbar) / standardDev)
    Next
    Kurtosis = sum / (pos) - 3
    
End Function

Public Function StudentsT(b As DataStats, ByRef prob As Double) As Double
'
' Student's t-test fo signifcantly different means
' Call CalcStats() on both Data's before calling this member
'
    Dim var1 As Double
    Dim var2 As Double
    Dim svar As Double
    Dim df As Long
    Dim t As Double
    
    var1 = aveVar
    var2 = b.aveVar
    
    df = Count + b.Count - 2
    
    svar = ((Count - 1) * var1 + (b.Count - 1) * var2) / df
    
    t = (Mean - b.Mean) / Sqr(svar * (1# / Count + 1# / b.Count))
    prob = betai(0.5 * df, 0.5, df / (df + t * t))
    
    StudentsT = t
End Function

Friend Function aveVar() As Double
    Dim ep As Double
    Dim s As Double
    Dim var As Double
    
    Dim j As Long
    
    For j = 0 To pos - 1
        s = data(j) - xbar
        ep = ep + s
        var = var + s * s
    Next
    
    aveVar = (var - ep * ep / pos) / (pos - 1)
End Function

Private Function betai(a As Double, b As Double, x As Double) As Double
'
' returns the incomplete beta funciton Ix(a, b)
'
    Dim bt As Double
    
    If x < 0# Or x > 1# Then
        Err.Raise -50001, "Data", "Bad x in batai"
    End If
    
    If x = 0# Or x = 1# Then
        bt = 0#
    Else
        bt = Exp(gammaln(a + b) - gammaln(a) - gammaln(b) + a * Log(x) + b * Log(1# - x))
    End If
    
    If x < (a + 1#) / (a + b + 2#) Then
        betai = bt * betacf(a, b, x) / a
    Else
        betai = 1# - bt * betacf(b, a, 1# - x) / b
    End If
End Function

Public Function gammaln(xx As Double) As Double
'
' returns the value ln(tao(?)(xx)) for xx > 0
'
    Dim x As Double
    Dim y As Double
    Dim tmp As Double
    Dim ser As Double
    
    Dim j As Long
    
    Dim cof(0 To 5) As Double
    cof(0) = 76.1800917294715
    cof(1) = -86.5053203294168
    cof(2) = 24.0140982408309
    cof(3) = -1.23173957245015
    cof(4) = 1.20865097386618E-03
    cof(5) = -5.395239384953E-06
    
    y = xx
    x = xx
    tmp = x + 5.5
    tmp = tmp - (x + 0.5) * Log(tmp)
    ser = 1.00000000019001
    
    For j = 0 To 5
        y = y + 1
        ser = ser + cof(j) / y
    Next
    gammaln = -tmp + Log(2.506628274631 * ser / x)
    
End Function

Public Function gamma(d As Double) As Double
    gamma = Exp(gammaln(d))
End Function

Private Function betacf(a As Double, b As Double, x As Double) As Double
'
' used by betai:  evaluate continued fration for incomplete beta funciton by
' modified lentz's method (5.2)
'
    Dim m As Long
    Dim m2 As Long
    Dim aa As Double
    Dim c As Double
    Dim d As Double
    Dim del As Double
    Dim h As Double
    Dim qab As Double
    Dim qam As Double
    Dim qap As Double
    
    Const FPMIN As Double = 1E-30
    Const EPS As Double = 0.0000003
    Const MAXIT As Long = 200
    
    qab = a + b
    qap = a + 1#
    qam = a - 1#
    c = 1#
    d = 1# - qab * x / qap
    If Abs(d) < FPMIN Then
        d = FPMIN
    End If
    
    d = 1# / d
    h = d
    
    ' 100 is the maximum iterations
    For m = 1 To MAXIT
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then
            d = FPMIN
        End If
        c = 1# + aa / c
        If Abs(c) < FPMIN Then
            c = FPMIN
        End If
        d = 1# / d
        h = h * d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1# + aa * d
        If Abs(d) < FPMIN Then
            d = FPMIN
        End If
        c = 1# + aa / c
        If c < FPMIN Then
            c = FPMIN
        End If
        d = 1# / d
        del = d * c
        h = h * del
        If Abs(del - 1#) < EPS Then
            ' finished
            Exit For
        End If
    Next
    If m = MAXIT Then
        Err.Raise -43543, "betacf", "a or b too big, or MAXIt too small"
    End If
    betacf = h
End Function

Public Function FTest(b As DataStats, ByRef fstat As Double) As Double
'
' Small values of prob indicate that the two sets have significantly
' different variances.
'
' Returns the probability
'
    Dim df1 As Double
    Dim df2 As Double
    Dim prob As Double
    
    Dim avgVar1 As Double
    Dim avgVar2 As Double
    Dim ep As Double
    Dim s As Double
    
    Dim j As Long
    
    For j = 0 To Count - 1
        s = data(j) - xbar
        ep = ep + s
        avgVar1 = avgVar1 + s * s
    Next
    avgVar1 = (avgVar1 - ep * ep / Count) / (Count - 1)
    
    For j = 0 To b.Count - 1
        s = b.Item(j) - b.Mean
        ep = ep + s
        avgVar2 = avgVar2 + s * s
    Next
    avgVar2 = (avgVar2 - ep * ep / b.Count) / (b.Count - 1)
    
    If avgVar1 > avgVar2 Then
        If avgVar2 = 0 Then
            fstat = 0
        Else
            fstat = avgVar1 / avgVar2
        End If
        df1 = Count - 1
        df2 = b.Count - 1
    Else
        If avgVar1 = 0 Then
            fstat = 0
        Else
            fstat = avgVar2 / avgVar1
        End If
        df1 = b.Count - 1
        df2 = Count - 1
    End If
    
    prob = 2# * betai(0.5 * df2, 0.5 * df1, df2 / (df2 + df1 * fstat))
    If prob > 1# Then
        prob = 2# - prob
    End If
    
    Debug.Assert prob <= 1
    
    FTest = prob
End Function

Public Function ApproximateTTest(b As DataStats, ByVal prob, df As Double) As Double
'
'
'
    Dim var1 As Double
    Dim var2 As Double
    
    Dim t As Double
    
    'var1 = aveVar
    'var2 = B.aveVar
    
    var1 = Variance
    var2 = b.Variance
    
    t = (xbar - b.Mean) / Sqr(var1 / Count + var2 / b.Count)
    df = ((var1 / Count + var2 / b.Count) ^ 2) / (((var1 / Count) ^ 2) / (Count - 1) + ((var2 / b.Count) ^ 2) / (b.Count - 1))
    prob = betai(0.5 * df, 0.5 * df, df / (df + t * t))
    
    ApproximateTTest = t
End Function

Public Function OneSampleTTest(b As DataStats) As Double
'
'
'
    Dim t As Double
    If var = 0 Then
        ' I think this is wrong
        t = 4#
    Else
        t = (xbar - b.Mean) / (Sqr(var / Count))
    End If
    
    OneSampleTTest = t
End Function

Private Sub Class_Initialize()
    ReDim data(0 To 0)
    pos = 0
End Sub

#If 0 Then
'{This program performs the Brown-Forsythe & Levene Tests for homogeneity of
'variances for the two group case.
'Levene and Brown-Forsythe tests for homogeneity of variances (HOV)
'A important assumption in the t-test for mean differences is that the variances
'in the two groups are equal (homogeneous). Two powerful and commonly used tests
'of this assumption are the Levene test and the Brown-Forsythe modification of
'this test. However, it is important to realize that (1) the homogeneity of
'variances assumption is usually not as crucial as other assumptions for the
't-test for mean differences, in particular in the case of equal n, and (2) that
'the tests described below are not necessarily very robust themselves (e.g.,
'Glass and Hopkins, 1996, p. 436, call these tests "fatally flawed" see also the
'description of these tests below). If you are concerned about a violation of the
'HOV assumption, it is always advisable to repeat the key analyses using
'nonparametric methods.
'Levene 's test (homogeneity of variances): For each dependent variable, a t-test
'for mean differences is performed on the absolute deviations of values from the
'respective group means. If the Levene test is statistically significant, then
'the hypothesis of homogeneous variances should be rejected.
'Brown & Forsythe's test (homogeneity of variances): Recently, some authors
'(e.g., Glass and Hopkins, 1996) have called into question the power of the
'Levene test for unequal variances. Specifically, the absolute deviation (from
'the group means) scores can be expected to be highly skewed thus, the normality
'assumption for t-test for mean differences of those absolute deviation scores is
'usually violated. This poses a particular problem when there is unequal n in the
'two groups that are to be compared. A more robust test that is very similar to
'the Levene test has been proposed by Brown and Forsythe (1974). Instead of
'performing the t-test for mean differences on the deviations from the mean, one
'can perform the analysis on the deviations from the group medians. Olejnik and
'Algina (1987) have shown that this test will give quite accurate error rates
'even when the underlying distributions for the raw scores deviate significantly
'from the normal distribution. However, as Glass and Hopkins (1996, p. 436) have
'pointed out, both the Levene test as well as the Brown-Forsythe modification
'suffer from what those authors call a "fatal flaw," namely, that both tests
'themselves rely on the homogeneity of variances assumption (of the absolute
'deviations from the means or medians) and hence, it is not clear how robust
'these tests are themselves in the presence of significant variance heterogeneity
'and unequal n.
'Program written, modified, or edited at StatSoft, Inc.}

#End If

Public Function Levene(b As DataStats, prob) As Double
'
' large values of prob indiccate that the two sets have significantly different variances
'
     
    If Count = 0 Or b.Count = 0 Then
        Err.Raise 123, "StatPack", "No data in group"
    End If
       
    Dim z1() As Double
    ReDim z1(Count)
    
    Dim z2() As Double
    ReDim z2(b.Count)
    
    Dim med1 As Double
    Dim med2 As Double
    
    med1 = Median
    med2 = b.Median
    
    Dim j As Long
    Dim sum1 As Double
    Dim zmean1 As Double
    Dim zmean2 As Double
    
    sum1 = 0#
    For j = 0 To Count - 1
        z1(j) = Abs(data(j) - Mean)
        sum1 = sum1 + data(j)
    Next
    zmean1 = sum1 / Count
    
    sum1 = 0#
    For j = 0 To Count - 1
        z2(j) = Abs(b.Item(j))
        sum1 = sum1 + b.Item(j)
    Next
    zmean2 = sum1 / b.Count
    sum1 = 0#
    
    Dim sumsqr1 As Double
    Dim sumsqr2 As Double
    Dim sum2 As Double
    
    For j = 0 To Count - 1
        sum1 = sum1 + z1(j)
        sumsqr1 = sumsqr1 + (z1(j)) ^ 2
    Next
    
    For j = 0 To b.Count - 1
        sum2 = sum2 + z2(j)
        sumsqr2 = sumsqr2 + (z2(j)) ^ 2
    Next
    
    Dim sumsqr As Double
    Dim SST As Double
    Dim SSTRT As Double
    Dim SSE As Double
    Dim f As Double
    Dim pval As Double

    sumsqr = (sum1 + sum2) ^ 2
    SST = sumsqr1 + sumsqr2 - sumsqr / (Count + b.Count)
    SSTRT = (sum1 ^ 2) / Count + (sum2 ^ 2) / b.Count - sumsqr / (Count + b.Count)
    SSE = SST - SSTRT
    'degrees of freedom = Count + B.Count - 2
    f = SSTRT / (SSE / Count + b.Count - 2)
    pval = 1 - iFDistr(f, 1, Count + b.Count - 2)
    
    'Levene2 = pval > alpha
    prob = pval
    Levene = f
End Function

Public Function Median() As Double
    Dim da() As Double
    ReDim da(Count)
    
    Dim x As Long
    
    For x = 0 To Count - 1
        ArrayInsert da, data(x)
    Next
    
    Median = da(Count / 2)
End Function

Private Sub ArrayInsert(a() As Double, d As Double)
'
' insert n values into an n dimensional array
'
' called by:  RankSQV
'
    Dim x As Long
    
    For x = UBound(a) - 1 To LBound(a) Step -1
        If d > a(x) Then
            'move values down
            a(x + 1) = a(x)
        Else
            ' insert the value
            a(x + 1) = d
            Exit Sub
        End If
    Next
    ' d was greater than all the values in the array
    a(LBound(a)) = d
End Sub

Private Function iFDistr(x As Double, v As Double, w As Double) As Double
    iFDistr = (gamma((v + w) / 2)) / (gamma(v / 2) * gamma(w / 2)) * ((v / w) ^ (v / 2)) * (x ^ (v / 2)) * (1 + (v / w) * x) ^ (-(v + w) / 2)
End Function

Private Sub sort(N As Long, ar() As Double)
    Dim x As Long
    Dim y As Long
    
    For x = LBound(ar) To UBound(ar)
        For y = x To UBound(ar)
            If ar(x) > ar(y) Then
                Dim d As Double
                d = ar(x)
                ar(x) = ar(y)
                ar(y) = d
            End If
        Next
    Next
End Sub

Public Function WTestSmall(b As DataStats) As Double
'
'/*   13    Shapiro-Wilk W Test of Normality for Small Samples */
'   THIS IS THE SHAPIRO-WILK TEST FOR SMALL SAMPLES
'   THE SAMPLE SIZE MUST BE LESS THAN OR EQUAL TO 50
'   TEST13 SW(N)  =\002,f10.4
'

    '/* Local variables */
    Dim sumb As Double
    
    Dim sumx As Double
    Dim a(25) As Double
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim sumx2 As Double
    'Dim s2 As Double
    Dim x1() As Double
    
    Dim N As Long
    
    N = Count + b.Count
    ReDim x1(0 To N - 1)

    If (N > 50) Then
        Err.Raise 1234, "StatPack", "This version of the w-test only handles up to 50 samples"
        Exit Function
    End If
    
    If N >= 20 Then
        GoTo L53
    End If
    If N >= 10 Then
        GoTo L33
    End If
    If N <> 3 Then
        GoTo L21
    End If
    a(0) = 0.7071
    GoTo L115
L21:
    If N <> 4 Then
        GoTo L23
    End If
    a(0) = 0.6872
    a(1) = 0.1677
    GoTo L115
L23:
    If N <> 5 Then
        GoTo L25
    End If
    a(0) = 0.6646
    a(1) = 0.2413
    GoTo L115
L25:
    If N <> 6 Then
        GoTo L27
    End If
    a(0) = 0.6431
    a(1) = 0.2806
    a(2) = 0.0875
    GoTo L115
L27:
    If N <> 7 Then
        GoTo L29
    End If
    a(0) = 0.6233
    a(1) = 0.3031
    a(2) = 0.1401
    GoTo L115
L29:
    If N <> 8 Then
        GoTo L31
    End If
    a(0) = 0.6052
    a(1) = 0.3164
    a(2) = 0.1743
    a(3) = 0.0561
    GoTo L115
L31:
    If N <> 9 Then
        GoTo L33
    End If
    a(0) = 0.5888
    a(1) = 0.3244
    a(2) = 0.1976
    a(3) = 0.0947
    GoTo L115
L33:
    If N <> 10 Then
        GoTo L35
    End If
    a(0) = 0.5739
    a(1) = 0.3291
    a(2) = 0.2141
    a(3) = 0.1224
    a(4) = 0.0399
    GoTo L115
L35:
    If N <> 11 Then
    GoTo L37
    End If
    a(0) = 0.5601
    a(1) = 0.3315
    a(2) = 0.226
    a(3) = 0.1429
    a(4) = 0.0695
    GoTo L115
L37:
    If N <> 12 Then
        GoTo L39
    End If
    a(0) = 0.5475
    a(1) = 0.3325
    a(2) = 0.2347
    a(3) = 0.1586
    a(4) = 0.0922
    a(5) = 0.0303
    GoTo L115
L39:
    If N <> 13 Then
        GoTo L41
    End If
    a(0) = 0.5359
    a(1) = 0.3325
    a(2) = 0.2412
    a(3) = 0.1707
    a(4) = 0.1099
    a(5) = 0.0539
    GoTo L115
L41:
    If N <> 14 Then
        GoTo L43
    End If
    a(0) = 0.5251
    a(1) = 0.3318
    a(2) = 0.246
    a(3) = 0.1802
    a(4) = 0.124
    a(5) = 0.0727
    a(6) = 0.024
    GoTo L115
L43:
    If N <> 15 Then
        GoTo L45
    End If
    a(0) = 0.515
    a(1) = 0.3306
    a(2) = 0.2495
    a(3) = 0.1878
    a(4) = 0.1353
    a(5) = 0.088
    a(6) = 0.0433
    GoTo L115
L45:
    If N <> 16 Then
        GoTo L47
    End If
    a(0) = 0.5056
    a(1) = 0.329
    a(2) = 0.2521
    a(3) = 0.1939
    a(4) = 0.1447
    a(5) = 0.1005
    a(6) = 0.0593
    a(7) = 0.0196
    GoTo L115
L47:
    If N <> 17 Then
    GoTo L49
    End If
    a(0) = 0.4968
    a(1) = 0.3273
    a(2) = 0.254
    a(3) = 0.1988
    a(4) = 0.1524
    a(5) = 0.1109
    a(6) = 0.0725
    a(7) = 0.0359
    GoTo L115
L49:
    If N <> 18 Then
        GoTo L51
    End If
    a(0) = 0.4886
    a(1) = 0.3253
    a(2) = 0.2553
    a(3) = 0.2027
    a(4) = 0.1587
    a(5) = 0.1197
    a(6) = 0.0837
    a(7) = 0.0496
    a(8) = 0.0163
    GoTo L115
L51:
    If N <> 19 Then
        GoTo L53
    End If
    a(0) = 0.4808
    a(1) = 0.3232
    a(2) = 0.2561
    a(3) = 0.2059
    a(4) = 0.1641
    a(5) = 0.1271
    a(6) = 0.0932
    a(7) = 0.0612
    a(8) = 0.0303
    GoTo L115
L53:
    If N <> 20 Then
    GoTo L55
    End If
    a(0) = 0.4734
    a(1) = 0.3211
    a(2) = 0.2565
    a(3) = 0.2085
    a(4) = 0.1686
    a(5) = 0.1334
    a(6) = 0.1013
    a(7) = 0.0711
    a(8) = 0.0422
    a(9) = 0.014
    GoTo L115
L55:
    If N <> 21 Then
        GoTo L57
    End If
    a(0) = 0.4643
    a(1) = 0.3185
    a(2) = 0.2578
    a(3) = 0.2119
    a(4) = 0.1736
    a(5) = 0.1399
    a(6) = 0.1092
    a(7) = 0.0804
    a(8) = 0.053
    a(9) = 0.0263
    GoTo L115
L57:
    If N <> 22 Then
        GoTo L59
    End If
    a(0) = 0.459
    a(1) = 0.3156
    a(2) = 0.2571
    a(3) = 0.2131
    a(4) = 0.1764
    a(5) = 0.1443
    a(6) = 0.115
    a(7) = 0.0878
    a(8) = 0.0618
    a(9) = 0.0368
    a(10) = 0.0122
    GoTo L115
L59:
    If N <> 23 Then
        GoTo L61
    End If
    a(0) = 0.4542
    a(1) = 0.3126
    a(2) = 0.2563
    a(3) = 0.2139
    a(4) = 0.1787
    a(5) = 0.148
    a(6) = 0.1201
    a(7) = 0.0941
    a(8) = 0.0696
    a(9) = 0.0459
    a(10) = 0.0228
    GoTo L115
L61:
    If N <> 24 Then
        GoTo L63
    End If
    a(0) = 0.4493
    a(1) = 0.3098
    a(2) = 0.2554
    a(3) = 0.2145
    a(4) = 0.1807
    a(5) = 0.1512
    a(6) = 0.1245
    a(7) = 0.0997
    a(8) = 0.0764
    a(9) = 0.0539
    a(10) = 0.0321
    a(11) = 0.0107
    GoTo L115
L63:
    If N <> 25 Then
        GoTo L65
    End If
    a(0) = 0.445
    a(1) = 0.3069
    a(2) = 0.2543
    a(3) = 0.2148
    a(4) = 0.1822
    a(5) = 0.1539
    a(6) = 0.1283
    a(7) = 0.1046
    a(8) = 0.0823
    a(9) = 0.061
    a(10) = 0.0403
    a(11) = 0.02
    GoTo L115
L65:
    If N <> 26 Then
        GoTo L67
    End If
    a(0) = 0.4407
    a(1) = 0.3043
    a(2) = 0.2533
    a(3) = 0.2151
    a(4) = 0.1836
    a(5) = 0.1563
    a(6) = 0.1316
    a(7) = 0.1089
    a(8) = 0.0876
    a(9) = 0.0672
    a(10) = 0.0476
    a(11) = 0.0284
    a(12) = 0.0094
    GoTo L115
L67:
    If N <> 27 Then
        GoTo L69
    End If
    a(0) = 0.4366
    a(1) = 0.3018
    a(2) = 0.2522
    a(3) = 0.2152
    a(4) = 0.1848
    a(5) = 0.1584
    a(6) = 0.1346
    a(7) = 0.1128
    a(8) = 0.0923
    a(9) = 0.0728
    a(10) = 0.054
    a(11) = 0.0358
    a(12) = 0.0178
    GoTo L115
L69:
    If N <> 28 Then
        GoTo L71
    End If
    a(0) = 0.4328
    a(1) = 0.2992
    a(2) = 0.251
    a(3) = 0.2151
    a(4) = 0.1857
    a(5) = 0.1601
    a(6) = 0.1372
    a(7) = 0.1162
    a(8) = 0.0965
    a(9) = 0.0778
    a(10) = 0.0598
    a(11) = 0.0424
    a(12) = 0.0253
    a(13) = 0.0084
    GoTo L115
L71:
    If N <> 29 Then
        GoTo L73
    End If
    a(0) = 0.4291
    a(1) = 0.2968
    a(2) = 0.2499
    a(3) = 0.215
    a(4) = 0.1864
    a(5) = 0.1616
    a(6) = 0.1395
    a(7) = 0.1192
    a(8) = 0.1002
    a(9) = 0.0822
    a(10) = 0.065
    a(11) = 0.0483
    a(12) = 0.032
    a(13) = 0.0159
    GoTo L115
L73:
    If N <> 30 Then
        GoTo L75
    End If
    a(0) = 0.4254
    a(1) = 0.2944
    a(2) = 0.2487
    a(3) = 0.2148
    a(4) = 0.187
    a(5) = 0.163
    a(6) = 0.1415
    a(7) = 0.1219
    a(8) = 0.1036
    a(9) = 0.0862
    a(10) = 0.0697
    a(11) = 0.0537
    a(12) = 0.0381
    a(13) = 0.0227
    a(14) = 0.0076
    GoTo L115
L75:
    If N <> 31 Then
        GoTo L77
    End If
    a(0) = 0.422
    a(1) = 0.2921
    a(2) = 0.2475
    a(3) = 0.2145
    a(4) = 0.1874
    a(5) = 0.1641
    a(6) = 0.1433
    a(7) = 0.1243
    a(8) = 0.1066
    a(9) = 0.0899
    a(10) = 0.0739
    a(11) = 0.0585
    a(12) = 0.0435
    a(13) = 0.0289
    a(14) = 0.0144
    GoTo L115
L77:
    If N <> 32 Then
        GoTo L79
    End If
    a(0) = 0.4188
    a(1) = 0.2898
    a(2) = 0.2463
    a(3) = 0.2141
    a(4) = 0.1878
    a(5) = 0.1651
    a(6) = 0.1449
    a(7) = 0.1265
    a(8) = 0.1093
    a(9) = 0.0931
    a(10) = 0.0777
    a(11) = 0.0629
    a(12) = 0.0485
    a(13) = 0.0344
    a(14) = 0.0206
    a(15) = 0.0068
    GoTo L115
L79:
    If N <> 33 Then
        GoTo L81
    End If
    a(0) = 0.4156
    a(1) = 0.2876
    a(2) = 0.2451
    a(3) = 0.2137
    a(4) = 0.188
    a(5) = 0.166
    a(6) = 0.1463
    a(7) = 0.1284
    a(8) = 0.1118
    a(9) = 0.0961
    a(10) = 0.0812
    a(11) = 0.0669
    a(12) = 0.053
    a(13) = 0.0395
    a(14) = 0.0262
    a(15) = 0.0131
    GoTo L115
L81:
    If N <> 34 Then
        GoTo L83
    End If
    a(0) = 0.4127
    a(1) = 0.2854
    a(2) = 0.2439
    a(3) = 0.2132
    a(4) = 0.1882
    a(5) = 0.1667
    a(6) = 0.1475
    a(7) = 0.1301
    a(8) = 0.114
    a(9) = 0.0988
    a(10) = 0.0844
    a(11) = 0.0706
    a(12) = 0.0572
    a(13) = 0.0441
    a(14) = 0.0314
    a(15) = 0.0187
    a(16) = 0.0062
    GoTo L115
L83:
    If N <> 35 Then
        GoTo L85
    End If
    a(0) = 0.4096
    a(1) = 0.2834
    a(2) = 0.2427
    a(3) = 0.2127
    a(4) = 0.1883
    a(5) = 0.1673
    a(6) = 0.1487
    a(7) = 0.1317
    a(8) = 0.116
    a(9) = 0.1013
    a(10) = 0.0873
    a(11) = 0.0739
    a(12) = 0.061
    a(13) = 0.0484
    a(14) = 0.0361
    a(15) = 0.0239
    a(16) = 0.0119
    GoTo L115
L85:
    If N <> 36 Then
        GoTo L87
    End If
    a(0) = 0.4068
    a(1) = 0.2813
    a(2) = 0.2415
    a(3) = 0.2121
    a(4) = 0.1883
    a(5) = 0.1678
    a(6) = 0.1496
    a(7) = 0.1331
    a(8) = 0.1179
    a(9) = 0.1036
    a(10) = 0.09
    a(11) = 0.077
    a(12) = 0.0645
    a(13) = 0.0523
    a(14) = 0.0404
    a(15) = 0.0287
    a(16) = 0.0172
    a(17) = 0.0057
    GoTo L115
L87:
    If N <> 37 Then
        GoTo L89
    End If
    a(0) = 0.404
    a(1) = 0.2794
    a(2) = 0.2403
    a(3) = 0.2116
    a(4) = 0.1883
    a(5) = 0.1683
    a(6) = 0.1505
    a(7) = 0.1344
    a(8) = 0.1196
    a(9) = 0.1056
    a(10) = 0.0924
    a(11) = 0.0798
    a(12) = 0.0677
    a(13) = 0.0559
    a(14) = 0.0444
    a(15) = 0.0331
    a(16) = 0.022
    a(17) = 0.011
    GoTo L115
L89:
    If N <> 38 Then
        GoTo L91
    End If
    a(0) = 0.4015
    a(1) = 0.2774
    a(2) = 0.2391
    a(3) = 0.211
    a(4) = 0.1881
    a(5) = 0.1686
    a(6) = 0.1513
    a(7) = 0.1356
    a(8) = 0.1211
    a(9) = 0.1075
    a(10) = 0.0947
    a(11) = 0.0824
    a(12) = 0.0706
    a(13) = 0.0592
    a(14) = 0.0481
    a(15) = 0.0372
    a(16) = 0.0264
    a(17) = 0.0158
    a(18) = 0.0053
    GoTo L115
L91:
    If N <> 39 Then
        GoTo L93
    End If
    a(0) = 0.3989
    a(1) = 0.2755
    a(2) = 0.238
    a(3) = 0.2104
    a(4) = 0.188
    a(5) = 0.1689
    a(6) = 0.152
    a(7) = 0.1366
    a(8) = 0.1225
    a(9) = 0.1092
    a(10) = 0.0967
    a(11) = 0.0848
    a(12) = 0.0733
    a(13) = 0.0622
    a(14) = 0.0515
    a(15) = 0.0409
    a(16) = 0.0305
    a(17) = 0.0203
    a(18) = 0.0101
    GoTo L115
L93:
    If N <> 40 Then
        GoTo L95
    End If
    a(0) = 0.3964
    a(1) = 0.2737
    a(2) = 0.2368
    a(3) = 0.2098
    a(4) = 0.1878
    a(5) = 0.1691
    a(6) = 0.1526
    a(7) = 0.1376
    a(8) = 0.1237
    a(9) = 0.1108
    a(10) = 0.0986
    a(11) = 0.087
    a(12) = 0.0759
    a(13) = 0.0651
    a(14) = 0.0546
    a(15) = 0.0444
    a(16) = 0.0343
    a(17) = 0.0244
    a(18) = 0.0146
    a(19) = 0.0049
    GoTo L115
L95:
    If N <> 41 Then
        GoTo L97
    End If
    a(0) = 0.394
    a(1) = 0.2719
    a(2) = 0.2357
    a(3) = 0.2091
    a(4) = 0.1876
    a(5) = 0.1693
    a(6) = 0.1531
    a(7) = 0.1384
    a(8) = 0.1249
    a(9) = 0.1123
    a(10) = 0.1004
    a(11) = 0.0891
    a(12) = 0.0782
    a(13) = 0.0677
    a(14) = 0.0575
    a(15) = 0.0476
    a(16) = 0.0379
    a(17) = 0.0283
    a(18) = 0.0188
    a(19) = 0.0094
    GoTo L115
L97:
    If N <> 42 Then
        GoTo L99
    End If
    a(0) = 0.3917
    a(1) = 0.2701
    a(2) = 0.2345
    a(3) = 0.2085
    a(4) = 0.1874
    a(5) = 0.1694
    a(6) = 0.1535
    a(7) = 0.1392
    a(8) = 0.1259
    a(9) = 0.1136
    a(10) = 0.102
    a(11) = 0.0909
    a(12) = 0.0804
    a(13) = 0.0701
    a(14) = 0.0602
    a(15) = 0.0506
    a(16) = 0.0411
    a(17) = 0.0318
    a(18) = 0.0227
    a(19) = 0.0136
    a(20) = 0.0045
    GoTo L115
L99:
    If N <> 43 Then
        GoTo L101
    End If
    a(0) = 0.3894
    a(1) = 0.2684
    a(2) = 0.2334
    a(3) = 0.2078
    a(4) = 0.1871
    a(5) = 0.1695
    a(6) = 0.1539
    a(7) = 0.1398
    a(8) = 0.1269
    a(9) = 0.1149
    a(10) = 0.1035
    a(11) = 0.0927
    a(12) = 0.0824
    a(13) = 0.0724
    a(14) = 0.0628
    a(15) = 0.0534
    a(16) = 0.0442
    a(17) = 0.0352
    a(18) = 0.0263
    a(19) = 0.0175
    a(20) = 0.0087
    GoTo L115
L101:
    If N <> 44 Then
        GoTo L103
    End If
    a(0) = 0.3872
    a(1) = 0.2667
    a(2) = 0.2323
    a(3) = 0.2072
    a(4) = 0.1868
    a(5) = 0.1695
    a(6) = 0.1542
    a(7) = 0.1405
    a(8) = 0.1278
    a(9) = 0.116
    a(10) = 0.1049
    a(11) = 0.0943
    a(12) = 0.0842
    a(13) = 0.0745
    a(14) = 0.0651
    a(15) = 0.056
    a(16) = 0.0471
    a(17) = 0.0383
    a(18) = 0.0296
    a(19) = 0.0211
    a(20) = 0.0126
    a(21) = 0.0042
    GoTo L115
L103:
    If N <> 45 Then
        GoTo L105
    End If
    a(0) = 0.385
    a(1) = 0.2651
    a(2) = 0.2313
    a(3) = 0.2065
    a(4) = 0.1865
    a(5) = 0.1695
    a(6) = 0.1545
    a(7) = 0.141
    a(8) = 0.1286
    a(9) = 0.117
    a(10) = 0.1062
    a(11) = 0.0959
    a(12) = 0.086
    a(13) = 0.0765
    a(14) = 0.0673
    a(15) = 0.0584
    a(16) = 0.0497
    a(17) = 0.0412
    a(18) = 0.0328
    a(19) = 0.0245
    a(20) = 0.0163
    a(21) = 0.0081
    GoTo L115
L105:
    If N <> 46 Then
        GoTo L107
    End If
    a(0) = 0.383
    a(1) = 0.2635
    a(2) = 0.2302
    a(3) = 0.2058
    a(4) = 0.1862
    a(5) = 0.1695
    a(6) = 0.1548
    a(7) = 0.1415
    a(8) = 0.1293
    a(9) = 0.118
    a(10) = 0.1073
    a(11) = 0.0972
    a(12) = 0.0876
    a(13) = 0.0783
    a(14) = 0.0694
    a(15) = 0.0607
    a(16) = 0.0522
    a(17) = 0.0439
    a(18) = 0.0357
    a(19) = 0.0277
    a(20) = 0.0197
    a(21) = 0.0118
    a(22) = 0.0039
    GoTo L115
L107:
    If N <> 47 Then
        GoTo L109
    End If
    a(0) = 0.3808
    a(1) = 0.262
    a(2) = 0.2291
    a(3) = 0.2052
    a(4) = 0.1859
    a(5) = 0.1695
    a(6) = 0.155
    a(7) = 0.142
    a(8) = 0.13
    a(9) = 0.1189
    a(10) = 0.1085
    a(11) = 0.0986
    a(12) = 0.0892
    a(13) = 0.0801
    a(14) = 0.0713
    a(15) = 0.0628
    a(16) = 0.0546
    a(17) = 0.0465
    a(18) = 0.0385
    a(19) = 0.0307
    a(20) = 0.0229
    a(21) = 0.0153
    a(22) = 0.0076
    GoTo L115
L109:
    If N <> 48 Then
        GoTo L111
    End If
    a(0) = 0.3789
    a(1) = 0.2604
    a(2) = 0.2281
    a(3) = 0.2045
    a(4) = 0.1855
    a(5) = 0.1693
    a(6) = 0.1551
    a(7) = 0.1423
    a(8) = 0.1306
    a(9) = 0.1197
    a(10) = 0.1095
    a(11) = 0.0998
    a(12) = 0.0906
    a(13) = 0.0817
    a(14) = 0.0731
    a(15) = 0.0648
    a(16) = 0.0568
    a(17) = 0.0489
    a(18) = 0.0411
    a(19) = 0.0335
    a(20) = 0.0259
    a(21) = 0.0185
    a(22) = 0.0111
    a(23) = 0.0037
    GoTo L115
L111:
    If N <> 49 Then
        GoTo L113
    End If
    a(0) = 0.377
    a(1) = 0.2589
    a(2) = 0.2271
    a(3) = 0.2038
    a(4) = 0.1851
    a(5) = 0.1692
    a(6) = 0.1553
    a(7) = 0.1427
    a(8) = 0.1312
    a(9) = 0.1205
    a(10) = 0.1105
    a(11) = 0.101
    a(12) = 0.0919
    a(13) = 0.0832
    a(14) = 0.0748
    a(15) = 0.0667
    a(16) = 0.0588
    a(17) = 0.0511
    a(18) = 0.0436
    a(19) = 0.0361
    a(20) = 0.0288
    a(21) = 0.0215
    a(22) = 0.0143
    a(23) = 0.0071
    GoTo L115
L113:
    If N <> 50 Then
        GoTo L115
    End If
    a(0) = 0.3751
    a(1) = 0.2574
    a(2) = 0.226
    a(3) = 0.2032
    a(4) = 0.1847
    a(5) = 0.1691
    a(6) = 0.1554
    a(7) = 0.143
    a(8) = 0.1317
    a(9) = 0.1212
    a(10) = 0.1113
    a(11) = 0.102
    a(12) = 0.0932
    a(13) = 0.0846
    a(14) = 0.0764
    a(15) = 0.0685
    a(16) = 0.0608
    a(17) = 0.0532
    a(18) = 0.0459
    a(19) = 0.0386
    a(20) = 0.0314
    a(21) = 0.0244
    a(22) = 0.0174
    a(23) = 0.0104
    a(24) = 0.0035
L115:
    
    sumb = 0#
    sumx = 0#
    sumx2 = 0#
    
    k = N / 2
    
    ' put the data from both sets into an array
    'For i = 0 To Count - 1
    '    x1(i) = data(i)
    'Next
    
    'i = Count
    'For j = 0 To B.Count - 1
    '    x1(i) = B.Item(j)
    '    i = i + 1
    'Next

    'For i = 0 To N - 1
    '    sumx = sumx + x1(i)
    '    '/* Computing 2nd power */

    '    sumx2 = sumx2 + x1(i) * x1(i)
    '    a(i) = 0#
    'Next

    '/* Computing 2nd power */
    's2 = sumx2 - sumx * sumx / N
    
    'sort N, x1
    
    'For i = 0 To k
    '    j = N - (i + 1)
    '    sumb = sumb + a(i) * (x1(j) - x1(i))
    'Next

    '/* Computing 2nd power */
    'WTestSmall = sumb * sumb / s2

''''''''''''''''''''''''''''''''''''''''

    ' put the data from both sets into an array
    For i = 0 To Count - 1
        x1(i) = data(i) - xbar
    Next
    
    i = Count
    For j = 0 To b.Count - 1
        x1(i) = b.Item(j) - b.Mean
        i = i + 1
    Next
    
    sort N, x1
    
    For i = 0 To N - 1
        sumx = sumx + x1(i)
    Next
        
    sumx = sumx / N
    
    For i = 0 To N - 1
        sumx2 = sumx2 + (x1(i) - sumx) * (x1(i) - sumx)
    Next
    
    If ((N - 1) Mod 2) = 0 Then
        For i = 0 To k - 1
            sumb = sumb + a(i) * (x1(N - (i + 1)) - x1(i))
        Next
    Else
        For i = 0 To k
            sumb = sumb + a(i) * (x1(N - (i + 1)) - x1(i))
        Next
    End If
    
    If sumx2 = 0 Then
        WTestSmall = 0
    Else
        WTestSmall = sumb * sumb / sumx2
    End If
    
L160:

End Function

Public Function WTestCriticalValue(N As Long, alpha As Double) As Double
'
'  returns a value from this table
'
    Dim category As Long
    
    If alpha <= 0.01 Then
        category = 1
    ElseIf alpha <= 0.02 Then
        category = 2
    ElseIf alpha <= 0.05 Then
        category = 3
    ElseIf alpha <= 0.1 Then
        category = 4
    Else
        category = 5
    End If
    
    Dim table(3 To 50, 1 To 5) As Double
    
    'N , A0_01, A0_02, A0_05, A0_10, A0_50
    '3,
    table(3, 1) = 0.753: table(3, 2) = 0.756: table(3, 3) = 0.767: table(3, 4) = 0.789: table(3, 5) = 0.959
    '4,
    table(4, 1) = 0.687: table(4, 2) = 0.707: table(4, 3) = 0.748: table(4, 4) = 0.792: table(4, 5) = 0.935
    '5,
    table(5, 1) = 0.686: table(5, 2) = 0.715: table(5, 3) = 0.762: table(5, 4) = 0.806: table(5, 5) = 0.927
    '6,
    table(6, 1) = 0.713: table(6, 2) = 0.743: table(6, 3) = 0.788: table(6, 4) = 0.826: table(6, 5) = 0.927
    '7,
    table(7, 1) = 0.73: table(7, 2) = 0.76: table(7, 3) = 0.803: table(7, 4) = 0.838: table(7, 5) = 0.928
    '8,
    table(8, 1) = 0.749: table(8, 2) = 0.778: table(8, 3) = 0.818: table(8, 4) = 0.851: table(8, 5) = 0.932
    '9,
    table(9, 1) = 0.764: table(9, 2) = 0.791: table(9, 3) = 0.829: table(9, 4) = 0.859: table(9, 5) = 0.935
    '10,
    table(10, 1) = 0.781: table(10, 2) = 0.806: table(10, 3) = 0.842: table(10, 4) = 0.869: table(10, 5) = 0.938
    '11,
    table(11, 1) = 0.792: table(11, 2) = 0.817: table(11, 3) = 0.85: table(11, 4) = 0.876: table(11, 5) = 0.94
    '12,
    table(12, 1) = 0.805: table(12, 2) = 0.828: table(12, 3) = 0.859: table(12, 4) = 0.883: table(12, 5) = 0.943
    '13,
    table(13, 1) = 0.814: table(13, 2) = 0.837: table(13, 3) = 0.866: table(13, 4) = 0.889: table(13, 5) = 0.945
    '14,
    table(14, 1) = 0.825: table(14, 2) = 0.846: table(14, 3) = 0.874: table(14, 4) = 0.895: table(14, 5) = 0.947
    '15,
    table(15, 1) = 0.835: table(15, 2) = 0.855: table(15, 3) = 0.881: table(15, 4) = 0.901: table(15, 5) = 0.95
    '16,
    table(16, 1) = 0.844: table(16, 2) = 0.863: table(16, 3) = 0.887: table(16, 4) = 0.906: table(16, 5) = 0.952
    '17,
    table(17, 1) = 0.851: table(17, 2) = 0.869: table(17, 3) = 0.892: table(17, 4) = 0.91: table(17, 5) = 0.954
    '18,
    table(18, 1) = 0.858: table(18, 2) = 0.874: table(18, 3) = 0.897: table(18, 4) = 0.914: table(18, 5) = 0.956
    '19,
    table(19, 1) = 0.863: table(19, 2) = 0.879: table(19, 3) = 0.901: table(19, 4) = 0.917: table(19, 5) = 0.957
    '20,
    table(20, 1) = 0.868: table(20, 2) = 0.884: table(20, 3) = 0.905: table(20, 4) = 0.92: table(20, 5) = 0.959
    
    table(21, 1) = 0.873: table(21, 2) = 0.888: table(21, 3) = 0.908: table(21, 4) = 0.923: table(21, 5) = 0.96
    table(22, 1) = 0.878: table(22, 2) = 0.892: table(22, 3) = 0.911: table(22, 4) = 0.926: table(22, 5) = 0.961
    table(23, 1) = 0.881: table(23, 2) = 0.895: table(23, 3) = 0.914: table(23, 4) = 0.928: table(23, 5) = 0.962
    table(24, 1) = 0.884: table(24, 2) = 0.898: table(24, 3) = 0.916: table(24, 4) = 0.93: table(24, 5) = 0.963
    table(25, 1) = 0.888: table(25, 2) = 0.901: table(25, 3) = 0.918: table(25, 4) = 0.931: table(25, 5) = 0.964
                
    table(26, 1) = 0.891: table(26, 2) = 0.904: table(26, 3) = 0.92: table(26, 4) = 0.933: table(26, 5) = 0.965
    table(27, 1) = 0.894: table(27, 2) = 0.906: table(27, 3) = 0.923: table(27, 4) = 0.935: table(27, 5) = 0.965
    table(28, 1) = 0.896: table(28, 2) = 0.908: table(28, 3) = 0.924: table(28, 4) = 0.936: table(28, 5) = 0.966
    table(29, 1) = 0.898: table(29, 2) = 0.91: table(29, 3) = 0.926: table(29, 4) = 0.937: table(29, 5) = 0.966
    table(30, 1) = 0.9: table(30, 2) = 0.912: table(30, 3) = 0.927: table(30, 4) = 0.939: table(30, 5) = 0.967
                
    table(31, 1) = 0.902: table(31, 2) = 0.914: table(31, 3) = 0.929: table(31, 4) = 0.94: table(31, 5) = 0.967
    table(32, 1) = 0.904: table(32, 2) = 0.915: table(32, 3) = 0.93: table(32, 4) = 0.941: table(32, 5) = 0.968
    table(33, 1) = 0.906: table(33, 2) = 0.917: table(33, 3) = 0.931: table(33, 4) = 0.942: table(33, 5) = 0.968
    table(34, 1) = 0.908: table(34, 2) = 0.919: table(34, 3) = 0.933: table(34, 4) = 0.943: table(34, 5) = 0.969
    table(35, 1) = 0.91: table(35, 2) = 0.92: table(35, 3) = 0.934: table(35, 4) = 0.944: table(35, 5) = 0.969
                
    table(36, 1) = 0.912: table(36, 2) = 0.922: table(36, 3) = 0.935: table(36, 4) = 0.945: table(36, 5) = 0.97
    table(37, 1) = 0.914: table(37, 2) = 0.924: table(37, 3) = 0.936: table(37, 4) = 0.946: table(37, 5) = 0.97
    table(38, 1) = 0.916: table(38, 2) = 0.925: table(38, 3) = 0.938: table(38, 4) = 0.947: table(38, 5) = 0.971
    table(39, 1) = 0.917: table(39, 2) = 0.927: table(39, 3) = 0.939: table(39, 4) = 0.948: table(39, 5) = 0.971
    table(40, 1) = 0.919: table(40, 2) = 0.928: table(40, 3) = 0.94: table(40, 4) = 0.949: table(40, 5) = 0.972
                
    table(41, 1) = 0.92: table(41, 2) = 0.929: table(41, 3) = 0.941: table(41, 4) = 0.95: table(41, 5) = 0.972
    table(42, 1) = 0.922: table(42, 2) = 0.93: table(42, 3) = 0.942: table(42, 4) = 0.951: table(42, 5) = 0.972
    table(43, 1) = 0.923: table(43, 2) = 0.932: table(43, 3) = 0.943: table(43, 4) = 0.951: table(43, 5) = 0.973
    table(44, 1) = 0.924: table(44, 2) = 0.933: table(44, 3) = 0.944: table(44, 4) = 0.952: table(44, 5) = 0.973
    table(45, 1) = 0.926: table(45, 2) = 0.934: table(45, 3) = 0.945: table(45, 4) = 0.953: table(45, 5) = 0.973
                
    table(46, 1) = 0.927: table(46, 2) = 0.935: table(46, 3) = 0.945: table(46, 4) = 0.953: table(46, 5) = 0.974
    table(47, 1) = 0.928: table(47, 2) = 0.936: table(47, 3) = 0.946: table(47, 4) = 0.954: table(47, 5) = 0.974
    table(48, 1) = 0.929: table(48, 2) = 0.937: table(48, 3) = 0.947: table(48, 4) = 0.954: table(48, 5) = 0.974
    table(49, 1) = 0.929: table(49, 2) = 0.937: table(49, 3) = 0.947: table(49, 4) = 0.955: table(49, 5) = 0.974
    table(50, 1) = 0.93: table(50, 2) = 0.938: table(50, 3) = 0.947: table(50, 4) = 0.955: table(50, 5) = 0.974

    WTestCriticalValue = table(N, category)
End Function

Public Function TCriticalValue2(N As Long, alpha As Double) As Double
'
' I think this table is incorrect, but it is required by Levene's to
' match the results of Biostat.
'
    Dim table(1 To 50, 0 To 3) As Double
    Dim category As Long
    
    If alpha <= 0.05 Then
        category = 3
    ElseIf alpha <= 0.1 Then
        category = 2
    ElseIf alpha <= 0.2 Then
        category = 1
    Else
        category = 0
    End If
    
    'df , A0_25, A0_20, A0_10, A_05
    '1,
    table(1, 0) = 1.376: table(1, 1) = 3: table(1, 2) = 6.31: table(1, 3) = 12.796
    '2,
    table(2, 0) = 1.061: table(2, 1) = 2: table(2, 2) = 2.92: table(2, 3) = 4.303
    '3,
    table(3, 0) = 0.978: table(3, 1) = 2: table(3, 2) = 2.353: table(3, 3) = 3.182
    '4,
    table(4, 0) = 0.941: table(4, 1) = 2: table(4, 2) = 2.132: table(4, 3) = 2.776
    '5,
    table(5, 0) = 0.92: table(5, 1) = 1: table(5, 2) = 2.015: table(5, 3) = 2.571
    '6,
    table(6, 0) = 0.906: table(6, 1) = 1: table(6, 2) = 1.943: table(6, 3) = 2.447
    '7,
    table(7, 0) = 0.896: table(7, 1) = 1: table(7, 2) = 1.895: table(7, 3) = 2.365
    '8,
    table(8, 0) = 0.889: table(8, 1) = 1: table(8, 2) = 1.86: table(8, 3) = 2.306
    '9,
    table(9, 0) = 0.883: table(9, 1) = 1: table(9, 2) = 1.833: table(9, 3) = 2.262
    '10,
    table(10, 0) = 0.879: table(10, 1) = 1: table(10, 2) = 1.812: table(10, 3) = 2.228
    '11,
    table(11, 0) = 0.876: table(11, 1) = 1: table(11, 2) = 1.796: table(11, 3) = 2.201
    '12,
    table(12, 0) = 0.873: table(12, 1) = 1: table(12, 2) = 1.782: table(12, 3) = 2.179
    '13,
    table(13, 0) = 0.87: table(13, 1) = 1: table(13, 2) = 1.771: table(13, 3) = 2.16
    '14,
    table(14, 0) = 0.868: table(14, 1) = 1: table(14, 2) = 1.761: table(14, 3) = 2.145
    '15,
    table(15, 0) = 0.866: table(15, 1) = 1: table(15, 2) = 1.753: table(15, 3) = 2.131
    '16,
    table(16, 0) = 0.865: table(16, 1) = 1: table(16, 2) = 1.746: table(16, 3) = 2.12
    '17,
    table(17, 0) = 0.863: table(17, 1) = 1: table(17, 2) = 1.74: table(17, 3) = 2.11
    '18,
    table(18, 0) = 0.862: table(18, 1) = 1: table(18, 2) = 1.734: table(18, 3) = 2.101
    '19,
    table(19, 0) = 0.861: table(19, 1) = 1: table(19, 2) = 1.729: table(19, 3) = 2.093
    '20,
    table(20, 0) = 0.86: table(20, 1) = 1: table(20, 2) = 1.725: table(20, 3) = 2.086
    '21,
    table(21, 0) = 0.859: table(21, 1) = 1: table(21, 2) = 1.721: table(21, 3) = 2.08
    '22,
    table(22, 0) = 0.858: table(22, 1) = 1: table(22, 2) = 1.717: table(22, 3) = 2.074
    '23,
    table(23, 0) = 0.858: table(23, 1) = 1: table(23, 2) = 1.714: table(23, 3) = 2.069
    '24,
    table(24, 0) = 0.857: table(24, 1) = 1: table(24, 2) = 1.711: table(24, 3) = 2.064
    '25,
    table(25, 0) = 0.856: table(25, 1) = 1: table(25, 2) = 1.708: table(25, 3) = 2.06
'????????
    '26
    table(26, 0) = 0.856: table(26, 1) = 1: table(26, 2) = 1.71: table(26, 3) = 2.06
    ' 27
    table(27, 0) = 0.855: table(27, 1) = 1: table(27, 2) = 1.7: table(27, 3) = 2.05
    ' 28
    table(28, 0) = 0.855: table(28, 1) = 1: table(28, 2) = 1.7: table(28, 3) = 2.05
    ' 29
    table(29, 0) = 0.854: table(29, 1) = 1: table(29, 2) = 1.7: table(29, 3) = 2.04
    ' 30
    table(30, 0) = 0.854: table(30, 1) = 1: table(30, 2) = 1.7: table(30, 3) = 2.04
    
    Dim x As Long
    For x = 31 To 40
        table(x, 0) = 0.851: table(x, 1) = 1: table(x, 2) = 1.68: table(x, 3) = 2.02
    Next
    For x = 41 To 50
        table(x, 0) = 0.848: table(x, 1) = 1: table(x, 2) = 1.67: table(x, 3) = 2#
    Next
    
    TCriticalValue2 = table(N, category)
End Function

Public Function TCriticalValue(N As Long, alpha As Double) As Double
'
' Do a table lookup of t critical values.
'
    Dim table(1 To 50, 0 To 3) As Double
    Dim category As Long
    
    ' get a table column for the alpha level
    If alpha <= 0.025 Then
        category = 3
    ElseIf alpha <= 0.05 Then
        category = 2
    ElseIf alpha <= 0.1 Then
        category = 1
    Else
        category = 0
    End If
        
    'df , A0_20, A0_10, A0_05, A_025
    '1,
    table(1, 0) = 1.376: table(1, 1) = 3.08: table(1, 2) = 6.31: table(1, 3) = 12.796
    '2,
    table(2, 0) = 1.061: table(2, 1) = 1.89: table(2, 2) = 2.92: table(2, 3) = 4.303
    '3,
    table(3, 0) = 0.978: table(3, 1) = 1.64: table(3, 2) = 2.353: table(3, 3) = 3.182
    '4,
    table(4, 0) = 0.941: table(4, 1) = 1.53: table(4, 2) = 2.132: table(4, 3) = 2.776
    '5,
    table(5, 0) = 0.92: table(5, 1) = 1.48: table(5, 2) = 2.015: table(5, 3) = 2.571
    '6,
    table(6, 0) = 0.906: table(6, 1) = 1.44: table(6, 2) = 1.943: table(6, 3) = 2.447
    '7,
    table(7, 0) = 0.896: table(7, 1) = 1.42: table(7, 2) = 1.895: table(7, 3) = 2.365
    '8,
    table(8, 0) = 0.889: table(8, 1) = 1.4: table(8, 2) = 1.86: table(8, 3) = 2.306
    '9,
    table(9, 0) = 0.883: table(9, 1) = 1.38: table(9, 2) = 1.833: table(9, 3) = 2.262
    '10,
    table(10, 0) = 0.879: table(10, 1) = 1.37: table(10, 2) = 1.812: table(10, 3) = 2.228
    '11,
    table(11, 0) = 0.8: table(11, 1) = 1.36: table(11, 2) = 1.796: table(11, 3) = 2.201
    '12,
    table(12, 0) = 0.873: table(12, 1) = 1.36: table(12, 2) = 1.782: table(12, 3) = 2.179
    '13,
    table(13, 0) = 0.87: table(13, 1) = 1.35: table(13, 2) = 1.771: table(13, 3) = 2.16
    '14,
    table(14, 0) = 0.868: table(14, 1) = 1.34: table(14, 2) = 1.761: table(14, 3) = 2.145
    '15,
    table(15, 0) = 0.866: table(15, 1) = 1.34: table(15, 2) = 1.753: table(15, 3) = 2.131
    '16,
    table(16, 0) = 0.865: table(16, 1) = 1.34: table(16, 2) = 1.746: table(16, 3) = 2.12
    '17,
    table(17, 0) = 0.863: table(17, 1) = 1.33: table(17, 2) = 1.74: table(17, 3) = 2.11
    '18,
    table(18, 0) = 0.862: table(18, 1) = 1.33: table(18, 2) = 1.734: table(18, 3) = 2.101
    '19,
    table(19, 0) = 0.861: table(19, 1) = 1.33: table(19, 2) = 1.729: table(19, 3) = 2.093
    '20,
    table(20, 0) = 0.86: table(20, 1) = 1.32: table(20, 2) = 1.725: table(20, 3) = 2.086
    '21,
    table(21, 0) = 0.859: table(21, 1) = 1.32: table(21, 2) = 1.721: table(21, 3) = 2.08
    '22,
    table(22, 0) = 0.858: table(22, 1) = 1.32: table(22, 2) = 1.717: table(22, 3) = 2.074
    '23,
    table(23, 0) = 0.858: table(23, 1) = 1.32: table(23, 2) = 1.714: table(23, 3) = 2.069
    '24,
    table(24, 0) = 0.857: table(24, 1) = 1.32: table(24, 2) = 1.711: table(24, 3) = 2.064
    '25,
    table(25, 0) = 0.856: table(25, 1) = 1.32: table(25, 2) = 1.708: table(25, 3) = 2.06

    '26
    table(26, 0) = 0.856: table(26, 1) = 1: table(26, 2) = 1.71: table(26, 3) = 2.06
    ' 27
    table(27, 0) = 0.855: table(27, 1) = 1: table(27, 2) = 1.7: table(27, 3) = 2.05
    ' 28
    table(28, 0) = 0.855: table(28, 1) = 1: table(28, 2) = 1.7: table(28, 3) = 2.05
    ' 29
    table(29, 0) = 0.854: table(29, 1) = 1: table(29, 2) = 1.7: table(29, 3) = 2.04
    ' 30
    table(30, 0) = 0.854: table(30, 1) = 1: table(30, 2) = 1.7: table(30, 3) = 2.04
    
    Dim x As Long
    For x = 31 To 40
        table(x, 0) = 0.851: table(x, 1) = 1: table(x, 2) = 1.68: table(x, 3) = 2.02
    Next
    For x = 41 To 50
        table(x, 0) = 0.848: table(x, 1) = 1: table(x, 2) = 1.67: table(x, 3) = 2#
    Next
    
    TCriticalValue = table(N, category)
End Function

Public Function LeveneBioStatWay(b As DataStats) As Double
'
'  This is the alternate x2 < x1 way.
'
    ' 1. the absolute value of the difference between each replicate and the
    '    mean for its repective treatment is calculated.
    
    ' 2.  A 2-tailed t-test is performed between the test treatment absolute
    '     differences and the reference treatment absolute differences.
    
    Dim ds1 As DataStats
    Dim ds2 As DataStats
    
    Set ds1 = New DataStats
    Set ds2 = New DataStats
    
    Dim i As Long
    
    For i = 0 To Count - 1
        ds1.AppendData Abs(data(i) - Mean)
    Next
    ds1.CalcStats
    
    For i = 0 To b.Count - 1
        ds2.AppendData Abs(b.Item(i) - b.Mean)
    Next
    ds2.CalcStats
    
    Dim t As Double
    Dim tprob As Double
    
    t = ds1.StudentsT(ds2, tprob)
    
    LeveneBioStatWay = Abs(t)
    
End Function

Public Function UStat(b As DataStats) As Double
'
' mann-whitney
'
    Dim N As Double
    Dim i As Long
    Dim R1 As Long
    
    N = Count + b.Count
        
    R1 = SortSumR1(b)
    
    UStat = Count * b.Count + (Count * (Count + 1)) / 2 - R1
End Function

Public Function UStatAlt(b As DataStats) As Double
'
' mann-whitney
'
    Dim N As Double
    Dim i As Long
    Dim R1 As Long
    
    N = Count + b.Count
        
    R1 = b.SortSumR1(Me)
    
    UStatAlt = Count * b.Count + (b.Count * (b.Count + 1)) / 2 - R1
End Function

Private Sub ArrayPairInsert(a() As pair, d As Double, one As Boolean)
'
' insert n value into an n dimensional array
'
' called by:  RankSQV
'
    Dim x As Long
    
    'For x = UBound(a) - 1 To LBound(a) Step -1
    '    If d > a(x).val Then
    '        'move values down
    '        a(x + 1).val = a(x).val
    '        a(x + 1).one = a(x).one
    '    Else
    '        ' insert the value
    '        a(x + 1).val = d
    '        a(x + 1).one = one
    '        Exit Sub
    '    End If
    'Next
    '' d was greater than all the values in the array
    'a(LBound(a)).val = d
    'a(LBound(a)).one = one
    
    For x = LBound(a) To UBound(a) - 1
        If d >= a(x).val Then
            '
            ' insert the value
            '
            Dim y As Long
            For y = UBound(a) To x + 1 Step -1
                a(y).val = a(y - 1).val
                a(y).one = a(y - 1).one
            Next
            a(x).val = d
            a(x).one = one
            Exit Sub
        End If
    Next
    ' shouldn't get here
    Debug.Assert False
End Sub


Public Function SortSumR1(b As DataStats) As Double
    Dim x As Long
    
    Dim sum As Long
    Dim N As Long
    Dim r As Double
    
    N = Count + b.Count
    
    Dim pairs() As pair
    ReDim pairs(N)
    
    For x = 0 To N
        pairs(x).val = -1
    Next
    
    For x = 0 To Count - 1
        ArrayPairInsert pairs, data(x), True
    Next
    
    For x = 0 To b.Count - 1
        ArrayPairInsert pairs, b.Item(x), False
    Next
    
    For x = 0 To N - 1
        If pairs(x).one = True Then
            r = r + AvgRank(pairs, x)
        End If
    Next
    
    SortSumR1 = r
End Function

Private Function AvgRank(pairs() As pair, ByVal x As Long)
    Dim sum As Double
    Dim N As Long
    Dim i As Long
    
    Dim target As Double
    
    target = pairs(x).val
    
    For i = 0 To UBound(pairs)
        If pairs(i).val = target Then
            ' average the ranks of the duplicates
            N = 1
            sum = i + 1
            Do While i < UBound(pairs)
                If Not (pairs(i).val = pairs(i + 1).val) Then
                    Exit Do
                End If
                
                N = N + 1
                i = i + 1
                sum = sum + i + 1
            Loop
            AvgRank = sum / N
            Exit Function
        End If
    Next
    
    Debug.Assert True
End Function

Public Function UCriticalPoint(N1 As Long, N2 As Long, alpha As Double) As Double

    Dim table As Collection
    Set table = New Collection
    
    'N1N2 , A0_10, A0_05
    table.Add Array(4, -1), "2,2" '"3,2",6, ' I'm not sure about this row
    
    table.Add Array(6, -1), "3,2" '"3,2",6,
    table.Add Array(8, 9), "3,3" ',8,9
    table.Add Array(19, 21), "3,8" ',19,21
    table.Add Array(8, -1), "4,2" ',8,
    table.Add Array(11, 12), "4,3" ', 11, 12
    table.Add Array(13, 15), "4,4" ',13,15
    table.Add Array(16, 18), "4,5"
    table.Add Array(9, 10), "5,2" ',9,10
    table.Add Array(13, 14), "5,3" ',13,14
    table.Add Array(16, 18), "5,4" ',16,18
    table.Add Array(20, 21), "5,5" ',20,21
    table.Add Array(23, 25), "5,6" ',23,25
    table.Add Array(11, 12), "6,2" ',11,12
    table.Add Array(15, 16), "6,3" ',15,16
    table.Add Array(19, 21), "6,4" ',19,21
    table.Add Array(23, 25), "6,5" ',23,25
    table.Add Array(27, 29), "6,6" ',27,29
    table.Add Array(13, 14), "7,2" ',13,14
    table.Add Array(17, 19), "7,3" ',17,19
    table.Add Array(22, 24), "7,4" ',22,24
    table.Add Array(27, 29), "7,5" ',27,29
    table.Add Array(31, 34), "7,6" ',31,34
    table.Add Array(36, 38), "7,7" ',36,38
    table.Add Array(14, 15), "8,2" ',14,15
    table.Add Array(19, 21), "8,3" ',19,21
    table.Add Array(25, 27), "8,4" ',25,27
    table.Add Array(30, 32), "8,5" ',30,32
    table.Add Array(35, 38), "8,6" ',35,38
    table.Add Array(40, 43), "8,7" ',40,43
    table.Add Array(45, 49), "8,8" ',45,49
    table.Add Array(9, -1), "9,1" ',9,
    table.Add Array(16, 17), "9,2" ',16,17
    table.Add Array(22, 23), "9,3" ',22,23
    table.Add Array(27, 30), "9,4" ',27,30
    table.Add Array(33, 36), "9,5" ',33,36
    table.Add Array(39, 42), "9,6" ',39,42
    table.Add Array(45, 48), "9,7" ',45,48
    table.Add Array(50, 54), "9,8" ',50,54
    table.Add Array(56, 60), "9,9" ',56,60
    table.Add Array(10, -1), "10,1" ',10,
    table.Add Array(17, 19), "10,2" ',17,19
    table.Add Array(24, 26), "10,3" ',24,26
    table.Add Array(30, 33), "10,4" ',30,33
    table.Add Array(37, 39), "10,5" ',37,39
    table.Add Array(43, 46), "10,6" ',43,46
    table.Add Array(49, 53), "10,7" ',49,53
    table.Add Array(56, 60), "10,8" ',56,60
    table.Add Array(62, 66), "10,9" ',62,66
    table.Add Array(68, 73), "10,10" ',68,73

    Dim pos As Long
    If alpha <= 0.05 Then
        pos = 1
    Else
        pos = 0
    End If
    
    Dim d As Double
    
    If N1 < N2 Then
        d = table(CStr(N2) & "," & CStr(N1))(pos)
    Else
        d = table(CStr(N1) & "," & CStr(N2))(pos)
    End If
    
    If d = -1 Then
        'Err.Raise -82483, "UCriticalPoint", "Invalid sample sizes"
    End If
    
    UCriticalPoint = d
End Function

Public Sub Rankit(b As DataStats, ByRef x1 As DataStats, ByRef x2 As DataStats)
'
' the transformed data will be returned in x1 and x2.  x1 = Me and x2 = B
'
    Dim N As Long
    
    N = Count + b.Count
    
    If N > 50 Then
        Err.Raise -3232, "Rankit", "Samples too large"
        Exit Sub
    End If
    
    Dim x As Long
      
    Dim pairs() As pair
    ReDim pairs(N)
    
    For x = 0 To N
        pairs(x).val = -1
    Next
    
    For x = 0 To Count - 1
        ArrayPairInsert pairs, data(x), True
    Next

'For x = 0 To N - 1
'    Debug.Print pairs(x).one
'Next
    For x = 0 To b.Count - 1
        ArrayPairInsert pairs, b.Item(x), False
    Next
    
'For x = 0 To N - 1
'    Debug.Print pairs(x).one
'Next
    
    ' reverse the array
    For x = 0 To N - 1
        If x >= N - x Then
            Exit For
        End If
        Dim v1 As Double
        Dim one1 As Boolean
        
        v1 = pairs(x).val
        one1 = pairs(x).one
        
        pairs(x).val = pairs((N - 1) - x).val
        pairs(x).one = pairs((N - 1) - x).one
        
        pairs((N - 1) - x).val = v1
        pairs((N - 1) - x).one = one1
    Next
    
'For x = 0 To N - 1
'    Debug.Print pairs(x).one
'Next
    '
    ' create new packs with the rankit scores as elements
    '
    For x = 0 To N - 1
        If pairs(x).one = True Then
            ' my data
            x1.AppendData AvgRankits(pairs, N, x)
        Else
            x2.AppendData AvgRankits(pairs, N, x)
        End If
    Next
    
    x1.CalcStats
    x2.CalcStats
        
End Sub

Private Function AvgRankits(pairs() As pair, N As Long, ByVal x As Long)
    Dim sum As Double
    Dim ncount As Long
    Dim i As Long
    
    Dim target As Double
    
    target = pairs(x).val
    
    For i = 0 To UBound(pairs)
        If pairs(i).val = target Then
            ' average the ranks of the duplicates
            ncount = 1
            sum = Rankits(N, i + 1)
            While (i < UBound(pairs) And pairs(i).val = pairs(i + 1).val)
                ncount = ncount + 1
                i = i + 1
                sum = sum + Rankits(N, i + 1)
            Wend
            AvgRankits = sum / ncount
            Exit Function
        End If
    Next
    
    Debug.Assert True
End Function


Private Function Rankits(N As Long, Rank As Long) As Double

    Dim table(3 To 20, 1 To 20) As Double

    'N , RANK1, RANK2, RANK3, RANK4, RANK5, RANK6, RANK7, RANK8, RANK9, RANK10, RANK11, RANK12, RANK13, RANK14, RANK15, RANK16, RANK17, RANK18, RANK19, RANK20
    '3,
    table(3, 1) = -0.846: table(3, 2) = 0#: table(3, 3) = 0.846
    table(4, 1) = -1.029: table(4, 2) = -0.297: table(4, 3) = 0.297: table(4, 4) = 1.029
    table(5, 1) = -1.163: table(5, 2) = -0.495: table(5, 3) = 0#: table(5, 4) = 0.495: table(5, 5) = 1.163
    table(6, 1) = -1.267: table(6, 2) = -0.642: table(6, 3) = -0.202: table(6, 4) = 0.202: table(6, 5) = 0.642: table(6, 6) = 1.267
    table(7, 1) = -1.352: table(7, 2) = -0.757: table(7, 3) = -0.353: table(7, 4) = 0#: table(7, 5) = 0.353: table(7, 6) = 0.757: table(7, 7) = 1.352
    table(8, 1) = -1.424: table(8, 2) = -0.852: table(8, 3) = -0.473: table(8, 4) = -0.153: table(8, 5) = 0.153: table(8, 6) = 0.473: table(8, 7) = 0.852: table(8, 8) = 1.424
    table(9, 1) = -1.485: table(9, 2) = -0.932: table(9, 3) = -0.572: table(9, 4) = -0.275: table(9, 5) = 0#: table(9, 6) = 0.275: table(9, 7) = 0.572: table(9, 8) = 0.932: table(9, 9) = 1.485
    table(10, 1) = -1.539: table(10, 2) = -1.001: table(10, 3) = -0.656: table(10, 4) = -0.376: table(10, 5) = -0.123: table(10, 6) = 0.123: table(10, 7) = 0.376: table(10, 8) = 0.656: table(10, 9) = 1.001: table(10, 10) = 1.539
    table(11, 1) = -1.586: table(11, 2) = -1.062: table(11, 3) = -0.729: table(11, 4) = -0.462: table(11, 5) = -0.225: table(11, 6) = 0#: table(11, 7) = 0.225: table(11, 8) = 0.462: table(11, 9) = 0.729: table(11, 10) = 1.062: table(11, 11) = 1.586
    table(12, 1) = -1.629: table(12, 2) = -1.116: table(12, 3) = -0.793: table(12, 4) = -0.537: table(12, 5) = -0.312: table(12, 6) = -0.103: table(12, 7) = 0.103: table(12, 8) = 0.312: table(12, 9) = 0.537: table(12, 10) = 0.793: table(12, 11) = 1.116: table(12, 12) = 1.629
    table(13, 1) = -1.668: table(13, 2) = -1.164: table(13, 3) = -0.85: table(13, 4) = -0.603: table(13, 5) = -0.388: table(13, 6) = -0.19: table(13, 7) = 0#: table(13, 8) = 0.19: table(13, 9) = 0.388: table(13, 10) = 0.603: table(13, 11) = 0.85: table(13, 12) = 1.164: table(13, 13) = 1.668
    table(14, 1) = -1.703: table(14, 2) = -1.208: table(14, 3) = -0.901: table(14, 4) = -0.662: table(14, 5) = -0.456: table(14, 6) = -0.267: table(14, 7) = -0.088: table(14, 8) = 0.088: table(14, 9) = 0.267: table(14, 10) = 0.456: table(14, 11) = 0.662: table(14, 12) = 0.901: table(14, 13) = 1.208: table(14, 14) = 1.703
    table(15, 1) = -1.736: table(15, 2) = -1.248: table(15, 3) = -0.948: table(15, 4) = -0.715: table(15, 5) = -0.516: table(15, 6) = -0.335: table(15, 7) = -0.165: table(15, 8) = 0#: table(15, 9) = 0.165: table(15, 10) = 0.335: table(15, 11) = 0.516: table(15, 12) = 0.715: table(15, 13) = 0.948: table(15, 14) = 1.248: table(15, 15) = 1.736
    table(16, 1) = -1.766: table(16, 2) = -1.285: table(16, 3) = -0.99: table(16, 4) = -0.763: table(16, 5) = -0.57: table(16, 6) = -0.396: table(16, 7) = -0.234: table(16, 8) = -0.077: table(16, 9) = 0.077: table(16, 10) = 0.234: table(16, 11) = 0.396: table(16, 12) = 0.57: table(16, 13) = 0.763: table(16, 14) = 0.99: table(16, 15) = 1.285: table(16, 16) = 1.766
    table(17, 1) = -1.794: table(17, 2) = -1.319: table(17, 3) = -1.029: table(17, 4) = -0.807: table(17, 5) = -0.619: table(17, 6) = -0.451: table(17, 7) = -0.295: table(17, 8) = -0.146: table(17, 9) = 0#: table(17, 10) = 0.146: table(17, 11) = 0.295: table(17, 12) = 0.451: table(17, 13) = 0.619: table(17, 14) = 0.807: table(17, 15) = 1.029: table(17, 16) = 1.319: table(17, 17) = 1.794
    table(18, 1) = -1.82: table(18, 2) = -1.35: table(18, 3) = -1.066: table(18, 4) = -0.848: table(18, 5) = -0.665: table(18, 6) = -0.502: table(18, 7) = -0.351: table(18, 8) = -0.208: table(18, 9) = -0.069: table(18, 10) = 0.069: table(18, 11) = 0.208: table(18, 12) = 0.351: table(18, 13) = 0.502: table(18, 14) = 0.665: table(18, 15) = 0.848: table(18, 16) = 1.066: table(18, 17) = 1.35: table(18, 18) = 1.82
    table(19, 1) = -1.844: table(19, 2) = -1.38: table(19, 3) = -1.099: table(19, 4) = -0.886: table(19, 5) = -0.707: table(19, 6) = -0.548: table(19, 7) = -0.402: table(19, 8) = -0.264: table(19, 9) = -0.131: table(19, 10) = 0#: table(19, 11) = 0.131: table(19, 12) = 0.264: table(19, 13) = 0.402: table(19, 14) = 0.548: table(19, 15) = 0.707: table(19, 16) = 0.886: table(19, 17) = 1.099: table(19, 18) = 1.38: table(19, 19) = 1.844
    '20,
    table(20, 1) = -1.867: table(20, 2) = -1.408: table(20, 3) = -1.131: table(20, 4) = -0.921: table(20, 5) = -0.745: table(20, 6) = -0.59: table(20, 7) = -0.448: table(20, 8) = -0.315: table(20, 9) = -0.187: table(20, 10) = -0.062: table(20, 11) = 0.062: table(20, 12) = 0.187: table(20, 13) = 0.315: table(20, 14) = 0.448: table(20, 15) = 0.59: table(20, 16) = 0.745: table(20, 17) = 0.921: table(20, 18) = 1.131: table(20, 19) = 1.408: table(20, 20) = 1.867

Debug.Assert table(N, Rank) <> 0 Or (N = 9 And Rank = 5) Or (N = 15 And Rank = 8) Or (N = 11 And Rank = 6) Or (N = 19 And Rank = 10) Or (N = 17 And Rank = 9) Or (N = 7 And Rank = 4)

    Rankits = table(N, Rank)
End Function

Public Function TransformASin() As DataStats
    Dim x As Long
    Dim x1 As DataStats
    Set x1 = New DataStats
    
    For x = 0 To Count - 1
        x1.AppendData ArcSin(Sqr(data(x)))
    Next
    x1.CalcStats
    
    Set TransformASin = x1
Bye:
    Exit Function

End Function

Public Function TransformSqrt375() As DataStats
    Dim x As Long
    Dim x1 As DataStats
    Set x1 = New DataStats
    
    For x = 0 To Count - 1
        x1.AppendData Sqr(data(x) + 0.375)
    Next
    x1.CalcStats
    
    Set TransformSqrt375 = x1
End Function

Public Function TransformLog10() As DataStats
    Dim x As Long
    Dim x1 As DataStats
    Set x1 = New DataStats
    
    For x = 0 To Count - 1
        x1.AppendData Log(data(x)) / Log(10) + 1
    Next
    x1.CalcStats
    
    Set TransformLog10 = x1
End Function

Private Function ArcSin(x As Double) As Double
    If x = 1 Then
        ArcSin = 1.570796327
    Else
        'ArcSin = Atn(x / ((1 - x ^ 2) ^ 0.5))
        ArcSin = Atn(x / (Sqr(1 - x ^ 2)))
    End If
End Function

Public Function ErrorFunction(x As Double)
    If x < 0# Then
        ErrorFunction = 1# + gammp(0.5, x * x)
    Else
        ErrorFunction = gammq(0.5, x * x)
    End If
End Function

Public Function gammp(a As Double, x As Double) As Double
    Dim gamser As Double
    Dim gammcf As Double
    
    If x < 0# Or a <= 0# Then
        Err.Raise -34323, "gammp", "bad, bad args"
    End If
    
    If x < (a + 1#) Then
        gamser = gser(a, x)
        gammp = gamser
        Exit Function
    Else
        gammcf = gcf(a, x)
        gammp = 1# - gammcf
        Exit Function
    End If
End Function

Private Function gammq(a As Double, x As Double) As Double
    
    Dim gamser As Double
    Dim gammcf As Double
    
    If x < 0# Or a <= 0# Then
        Err.Raise -34323, "gammp", "bad, bad args"
    End If
    
    If x < (a + 1#) Then
        gamser = gser(a, x)
        gammq = 1# - gamser
        Exit Function
    Else
        gammcf = gcf(a, x)
        gammq = gammcf
        Exit Function
    End If
End Function

Private Function gcf(a As Double, x As Double) As Double
    Dim i As Long
    Dim an As Double
    Dim b As Double
    Dim c As Double
    Dim d As Double
    Dim del As Double
    Dim h As Double
    Dim gln As Double
    
    gln = gammaln(a)
    b = x + 1# - a
    c = 1# / 1E-30
    d = 1# / b
    h = d
    For i = 1 To 200
        an = -i * (1 - a)
        b = b + 2#
        d = an * d + b
        If Abs(d) < 1E-30 Then
             d = 1E-30
        End If
        c = b + an / c
        If Abs(c) < 1E-30 Then
            c = 1E-30
        End If
        d = 1# / d
        del = d * c
        h = h * del
        If Abs(del) < 0.0000003 Then
            Exit For
        End If
    Next
    If i >= 200 Then
        Err.Raise -234324, "gcf", "failed"
    End If
    gcf = Exp(-x + a * Log(x) - gln) * h
End Function

Private Function gser(a As Double, x As Double) As Double
    Dim N As Long
    Dim sum As Double
    Dim del As Double
    Dim ap As Double
    Dim gln As Double
    
    gln = gammaln(a)
    If x <= 0# Then
        If x < 0# Then
            Err.Raise -234234, "gser", "x less than 0 in gser"
            Exit Function
        End If
        gser = 0#
        Exit Function
    End If
    
    ap = a
    sum = 1# / a
    del = sum
    
    For N = 1 To 200
        ap = ap + 1
        del = del * x / ap
        sum = sum + del
        If Abs(del) < Abs(sum) * 0.0000003 Then
            gser = sum * Exp(-x + a * Log(x) - gln)
            Exit Function
        End If
    Next
    
    Err.Raise -3432, "gser", "a too large or ITMaX to small"
End Function

Public Function StandardError(b As DataStats)
    StandardError = Sqr((standardDev * standardDev) / Count + (b.StandardDeviation * b.StandardDeviation) / b.Count)
End Function


Public Function toString() As String
    Dim x As Long
    Dim s As String
    
    For x = 0 To pos - 1
        s = s & CStr(x + 1) & ": " & CStr(data(x)) & Chr(13)
    Next
    s = s & "Mean: " & Mean & Chr(13)
    s = s & "STD: " & standardDev & Chr(13)
    
    toString = s
End Function
