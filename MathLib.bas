Attribute VB_Name = "MathLib"
'-------------------------------------------------------------------------------
' Excel Functions
'-------------------------------------------------------------------------------

Function LinearInterp(x As Object, y As Object, xnew As Double)
Dim obj As Object
Dim i As Integer
Dim x1 As Double
Dim x2 As Double
Dim y1 As Double
Dim y2 As Double
If x.Cells(1).Value > xnew Then
    Exit Function
End If

For i = 1 To x.Cells.Count
    If xnew < x.Cells(i).Value Then
        Exit For
    End If
Next i
x1 = x.Cells(i - 1).Value
x2 = x.Cells(i).Value
y1 = y.Cells(i - 1).Value
y2 = y.Cells(i).Value

LinearInterp = (xnew - x1) / (x2 - x1) * (y2 - y1) + y1
End Function

Function CubicSpline(x As Object, y As Object, xnew As Double) As Double
Dim N As Integer
Dim i As Integer
Dim xs() As Double
Dim ys() As Double

If x.Cells.Count <> y.Cells.Count Then
    Debug.Print "X and Y must be the same size"
    Exit Function
End If

If xnew < x.Cells(1).Value Then
    Debug.Print "Value out of range"
    Exit Function
End If
N = y.Cells.Count
If xnew > x.Cells(N).Value Then
    Debug.Print "value out of range"
    Exit Function
End If

ReDim xs(1 To N)
ReDim ys(1 To N)
For i = 1 To N
    xs(i) = x.Cells(i).Value
    ys(i) = y.Cells(i).Value
Next i
CubicSpline = func_cubic_spline(xs, ys, xnew)
End Function
'-------------------------------------------------------------------------------
' VBA Functions
'-------------------------------------------------------------------------------
Function func_cubic_spline(x() As Double, y() As Double, xnew As Double) As Double
Dim N As Integer
Dim i As Integer
Dim j As Integer
Dim dX() As Double
Dim dY() As Double
Dim a As Double
Dim b As Double
Dim k() As Double
Dim matA() As Double
Dim matB() As Double
Dim t As Double

dX = Diff(x)
dY = Diff(y)
N = UBound(x)
ReDim matA(1 To N, 1 To N)
ReDim matB(1 To N)

For i = 2 To N - 1
    matA(i, i - 1) = 1 / dX(i - 1)
    matA(i, i) = 2 / dX(i - 1) + 2 / dX(i)
    matA(i, i + 1) = 1 / dX(i)
    matB(i) = 3 * (dY(i - 1) / (dX(i - 1) ^ 2) + dY(i) / (dX(i) ^ 2))
Next i
matA(1, 1) = 2 / dX(1)
matA(1, 2) = 1 / dX(1)
matB(1) = 3 * dY(1) / (dX(1) ^ 2)
matA(N, N - 1) = 1 / dX(N - 1)
matA(N, N) = 2 / dX(N - 1)
matB(N) = 3 * (dY(N - 1) / (dX(N - 1) ^ 2))
k = LinearSolver(matA, matB)
For i = 1 To N
    If xnew < x(i) Then
        Exit For
    End If
Next i

t = (xnew - x(i - 1)) / dX(i - 1)
a = k(i - 1) * (dX(i - 1)) - (dY(i - 1))
b = -k(i) * dX(i - 1) + dY(i - 1)
func_cubic_spline = (1 - t) * y(i - 1) + t * y(i) + t * (1 - t) * ((1 - t) * a + t * b)
End Function

Function Diff(x() As Double) As Double()
    
    Dim N As Long
    Dim Df() As Double
    
    ReDim Df(LBound(x) To UBound(x) - 1)
    
    For N = LBound(x) To UBound(x) - 1
        Df(N) = x(N + 1) - x(N)
    Next N
    Diff = Df
End Function

Function LinearSolver(a() As Double, b() As Double) As Double()
Dim i As Integer
Dim j As Integer
Dim N As Integer
Dim m As Double
Dim x() As Double
If UBound(a, 1) <> UBound(a, 2) Then
    Debug.Print "The first matrix must be square"
    Exit Function
End If
If UBound(a, 1) <> UBound(b) Then
    Debug.Print "Size of A doesn't match with B"
    Exit Function
End If

N = UBound(b)
ReDim x(1 To N)

For i = 1 To N - 1
    For j = i + 1 To N
        If a(j, i) <> 0 And a(i, i) <> 0 Then
            m = -a(j, i) / a(i, i)
            Call RowOperation(a, j, i, m)
            b(j) = b(j) + m * b(i)
        End If
    Next j
Next i

For i = N To 2 Step -1
    For j = i - 1 To 1 Step -1
        If a(j, i) <> 0 And a(i, i) <> 0 Then
            m = -a(j, i) / a(i, i)
            Call RowOperation(a, j, i, m)
            b(j) = b(j) + m * b(i)
        End If
    Next j
Next i
For i = 1 To N
    x(i) = b(i) / a(i, i)
Next i
LinearSolver = x
End Function

Sub RowOperation(ByRef Mat() As Double, row1 As Integer, row2 As Integer, multiplier As Double)
Dim i As Integer
For i = 1 To UBound(Mat, 2)
    Mat(row1, i) = Mat(row1, i) + Mat(row2, i) * multiplier
Next i
End Sub
