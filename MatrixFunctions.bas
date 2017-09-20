Attribute VB_Name = "Module1"
    Dim a As Integer, b As Integer, c As Integer, d As Integer, i As Integer, j As Integer
    Dim matsol As Integer



Function addmatrix(mat1 As Variant, mat2 As Variant) As Variant
    Dim z() As Long
    
    
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   
   
   ReDim z(mat1row, mat1col)
   
   a = mat1row
   b = mat1col
   d = mat2col
   'l = mat1col
   
   i = 0
   j = 0
   'l = 0
    
    
    Do While i <= a                                 ''solution matrix
        Do While j <= b
            z(i, j) = mat1(i, j) + mat2(i, j)
            j = j + 1
        Loop
        i = i + 1
        j = 0
    Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'
    addmatrix = z
    


    End Function

Sub runaddmatrix()
    
    Dim mat1() As Long
    Dim mat2() As Long
    Dim z() As Long
    
    mat1row = 5
    m1col = 5
    mat1col = m1col
    mat2row = 5
    mat2col = 5

    
    ReDim mat1(mat1row, mat1col)
    ReDim mat2(mat2row, mat2col)
    
    
    
    a = mat1row
    b = mat1col
    c = mat2row
    d = mat2col
    
    ReDim z(mat1row, mat1col)
    
    i = 0
    j = 0
    
    If a = c Then                                   ''initializing mat1 matrix with random values
        If b = d Then
            Do While i <= a
                Do While j <= b
                    mat1(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
         Else: MsgBox "Matrix size mismatch error"
         End If
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    If a = c Then                                   ''initializing mat2 matrix with random values
        If b = d Then
            Do While i <= a
                Do While j <= b
                    mat2(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
         Else: MsgBox "Matrix size mismatch error"
         End If
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    
   
   ReDim z(mat1row, mat1col)
   
   z = addmatrix(mat1, mat2)
      
    ''print input and results to immediate window
'  For i = 0 To 5
'    Debug.Print z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab & "///"
'    Next
        
Debug.Print WriteArrayToImmediateWindow(mat1)
Debug.Print WriteArrayToImmediateWindow(mat2)
Debug.Print WriteArrayToImmediateWindow(z)

End Sub

Function subtmatrix(mat1 As Variant, mat2 As Variant) As Variant
    Dim z() As Long
    
    
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   
   
   ReDim z(mat1row, mat1col)
   
   a = mat1row
   b = mat1col
   d = mat2col
   'l = mat1col
   
   i = 0
   j = 0
   'l = 0
    
    
    Do While i <= a                                 ''solution matrix
        Do While j <= b
            z(i, j) = mat1(i, j) - mat2(i, j)
            j = j + 1
        Loop
        i = i + 1
        j = 0
    Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'
    subtmatrix = z
    


    End Function

Sub runsubtmatrix()
    
    Dim mat1() As Long
    Dim mat2() As Long
    Dim z() As Long
    
    mat1row = 5
    m1col = 5
    mat1col = m1col
    mat2row = 5
    mat2col = 5

    
    ReDim mat1(mat1row, mat1col)
    ReDim mat2(mat2row, mat2col)
    
    
    
    a = mat1row
    b = mat1col
    c = mat2row
    d = mat2col
    
    ReDim z(mat1row, mat1col)
    
    i = 0
    j = 0
    
    If a = c Then                                   ''initializing mat1 matrix with random values
        If b = d Then
            Do While i <= a
                Do While j <= b
                    mat1(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
         Else: MsgBox "Matrix size mismatch error"
         End If
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    If a = c Then                                   ''initializing mat2 matrix with random values
        If b = d Then
            Do While i <= a
                Do While j <= b
                    mat2(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
         Else: MsgBox "Matrix size mismatch error"
         End If
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    
   
   ReDim z(mat1row, mat1col)
   
   z = subtmatrix(mat1, mat2)
      
    ''print input and results to immediate window
'  For i = 0 To 5
'    Debug.Print z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab & "///"
'    Next
        
Debug.Print WriteArrayToImmediateWindow(mat1)
Debug.Print WriteArrayToImmediateWindow(mat2)
Debug.Print WriteArrayToImmediateWindow(z)

End Sub



Function multmatrix(mat1row As Integer, mat1col As Integer, mat2row As Integer, mat2col As Integer)
    Dim mat1() As Long
    Dim mat2() As Long
    Dim z() As Long
    
    ReDim mat1(mat1row, mat1col)
    ReDim mat2(mat2row, mat2col)
    
    
    a = mat1row
    b = mat1col
    c = mat2row
    d = mat2col
    
    
    
    i = 0
    j = 0
    
    If b = c Then                                   ''initializing mat1 matrix with random values
            Do While i <= a
                Do While j <= b
                    mat1(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    If b = c Then                                   ''initializing mat2 matrix with random values
            Do While i <= a
                Do While j <= b
                    mat2(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    Dim k As Integer, l As Integer, solrow As Integer, solcol As Integer
    k = 0
    l = 0
    solrow = a
    solcol = d
    
    ReDim z(solrow, solcol)
    
    Do While i <= a                                         ''i <= mat1row
        Do While j <= d                                     '' j <= mat2col
            Do While l <= b                                 '' l <= mat1col
            z(i, j) = z(i, j) + mat1(i, l) * mat2(l, j)
            l = l + 1
            Loop
        j = j + 1
        l = 0
        Loop
        i = i + 1
        j = 0
    Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'
    multmatrix = z
    


    End Function
    
     Sub runmultmatrix()
    
    mat1row = 1
    m1col = 3
    mat2row = 3
    mat2col = 1
'   Dim z() As Long
'   Dim iA As Long, jA As Long
'
'   ReDim z(mat1row, mat2col)
'
'   z = multmatrix(Int(mat1row), Int(m1col), Int(mat2row), Int(mat1col))
'
'    ''print input and results to immediate window
'
'    Debug.Print z(0, 0) & vbTab & z(0, 1) & vbTab & z(0, 2)
'    Debug.Print z(1, 0) & vbTab & z(1, 1) & vbTab & z(1, 2)
'    Debug.Print z(2, 0) & vbTab & z(2, 1) & vbTab & z(2, 2)
 
    Dim mat1() As Long
    Dim mat2() As Long
    Dim z() As Long
    
    ReDim mat1(mat1row, m1col)
    ReDim mat2(mat2row, mat2col)
    
    i = 0
    j = 0
    a = mat1row
    b = m1col
    c = mat2row
    d = mat2col
    
    If b = c Then                                   ''initializing mat1 matrix with random values
            Do While i <= a
                Do While j <= b
                    mat1(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    
    If b = c Then                                   ''initializing mat2 matrix with random values
            Do While i <= c
                Do While j <= d
                    mat2(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If
    
    i = 0
    j = 0
    Dim k As Integer, l As Integer, solrow As Integer, solcol As Integer
    k = 0
    l = 0
    solrow = a
    solcol = d
    

   
  ReDim z(a, d) As Long
   
   z = mult2matrix(mat1, mat2)


'Debug.Print z(0, 0)
'Debug.Print z(1, 0)


    Debug.Print WriteArrayToImmediateWindow(mat1)
    Debug.Print WriteArrayToImmediateWindow(mat2)
    Debug.Print WriteArrayToImmediateWindow(z)
    
End Sub



Function divmatrix(mat1 As Variant, mat2 As Variant) As Variant
   
   mat2row = UBound(mat2, 1)
   'mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
'
'   a = mat1row
'   d = mat2col
'   l = mat1col
'
   i = 0
   j = 0
'   l = 0
   
   Dim mat2inv() As Variant
   mat2inv = Application.MInverse(mat2)
   Dim m2invexp() As Variant
   ReDim m2invexp(0 To mat2row, 0 To mat2col)
   
            Do While i <= mat2row
                Do While j <= mat2col
                    m2invexp(i, j) = mat2inv(i + 1, j + 1)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'

   '     divmatrix = m2invexp
   divmatrix = multmatfordiv(mat1, m2invexp)
  '  Debug.Print WriteArraydiv(m2invexp)
   ' Debug.Print WriteArraydiv(mat2)
  '  Debug.Print WriteArraydiv(mat2inv)


    End Function
    
    
    
    Sub rundivmatrix()
    
    mat1row = 4
    m1col = 4
    mat2row = 4
    mat2col = 4
'   Dim z() As Long
'   Dim iA As Long, jA As Long
'
'   ReDim z(mat1row, mat2col)
'
'   z = multmatrix(Int(mat1row), Int(m1col), Int(mat2row), Int(mat1col))
'
'    ''print input and results to immediate window
'
'    Debug.Print z(0, 0) & vbTab & z(0, 1) & vbTab & z(0, 2)
'    Debug.Print z(1, 0) & vbTab & z(1, 1) & vbTab & z(1, 2)
'    Debug.Print z(2, 0) & vbTab & z(2, 1) & vbTab & z(2, 2)
 
    Dim mat1() As Variant
    Dim mat2() As Variant
    Dim z() As Variant
    
    ReDim mat1(mat1row, m1col) As Variant
    ReDim mat2(mat2row, mat2col) As Variant
    
    i = 0
    j = 0
    a = mat1row
    b = m1col
    c = mat2row
    d = mat2col

    If b = c Then                                   ''initializing mat1 matrix with random values
            Do While i <= a
                Do While j <= b
                    mat1(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If

    i = 0
    j = 0

    If b = c Then                                   ''initializing mat2 matrix with random values
            Do While i <= c
                Do While j <= d
                    mat2(i, j) = Int(Rnd * 100)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
     Else: MsgBox "Matrix size mismatch error"
    End If

    i = 0
    j = 0
    Dim k As Integer, l As Integer, solrow As Integer, solcol As Integer
    k = 0
    l = 0
    solrow = a
    solcol = d

'mat2(0, 0) = 4
'mat2(0, 1) = 7
'mat2(1, 0) = 2
'mat2(1, 1) = 6


  ReDim z(a, d)
   
   z = divmatrix(mat1, mat2)


'Debug.Print z(0, 0)
'Debug.Print z(1, 0)


    Debug.Print WriteArraydiv(mat1)
    Debug.Print WriteArraydiv(mat2)
    Debug.Print WriteArraydiv(z)
    
End Sub


Function multmatfordiv(mat1 As Variant, mat2 As Variant) As Variant
   
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   
   a = mat1row
   d = mat2col
   l = mat1col
   
   i = 0
   j = 0
   l = 0
   
   Dim z() As Variant
   ReDim z(a, d)
    
    Do While i <= a                                         ''i <= mat1row
        Do While j <= d                                     '' j <= mat2col
            Do While l <= b                                 '' l <= mat1col
            z(i, j) = z(i, j) + mat1(i, l) * mat2(l, j)
            l = l + 1
            Loop
        j = j + 1
        l = 0
        Loop
        i = i + 1
        j = 0
    Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'
   multmatfordiv = z
    


    End Function




    
    
    
''borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'' made some modifications to get it to display all arrays properly

Function WriteArrayToImmediateWindow(arrSubA As Variant)

Dim rowString As String
Dim iSubA As Long
Dim jSubA As Long

iSubA = 0
jSubA = 0

Dim rowmax As Integer, colmax As Integer
rowmax = UBound(arrSubA, 1)
colmax = UBound(arrSubA, 2)

rowString = ""

Debug.Print
Debug.Print
Debug.Print "The array is: "
For iSubA = 0 To rowmax
    rowString = arrSubA(iSubA, 0)
    For jSubA = 1 To colmax
        rowString = rowString & "," & arrSubA(iSubA, jSubA)
    Next jSubA
    Debug.Print rowString
Next iSubA

End Function
 
    
''borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'' made some modifications to get it to display all arrays properly

Function WriteArraydiv(arrSubA As Variant)

Dim rowString As String
Dim iSubA As Integer
Dim jSubA As Integer

iSubA = 0
jSubA = 0

Dim rowmax As Integer, colmax As Integer
rowmax = UBound(arrSubA, 1)
colmax = UBound(arrSubA, 2)

rowString = ""

Debug.Print
Debug.Print
Debug.Print "The array is: "
For iSubA = 0 To rowmax
    rowString = arrSubA(iSubA, 0)
    For jSubA = 1 To colmax
        rowString = rowString & "," & arrSubA(iSubA, jSubA)
    Next jSubA
    Debug.Print rowString
Next iSubA

End Function


Sub testdivmat()

        Dim amat() As Variant
        Dim amatinv() As Variant
        ReDim amat(1, 1)
        'ReDim amatinv(1, 1)
        amat(0, 0) = 4
        amat(0, 1) = 7
        amat(1, 0) = 2
        amat(1, 1) = 6
        'ReDim amatinv(1, 1)
        amatinv = divmatrix(amat, amat)
        
        
        Debug.Print WriteArraydiv(amatinv)
    End Sub
    
    
    Function mult2matrix(mat1 As Variant, mat2 As Variant) As Variant
   
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   
   a = mat1row
   d = mat2col
   l = mat1col
   
   i = 0
   j = 0
   l = 0
   
   Dim z() As Long
   ReDim z(a, d)
    
    Do While i <= a                                         ''i <= mat1row
        Do While j <= d                                     '' j <= mat2col
            Do While l <= b                                 '' l <= mat1col
            z(i, j) = z(i, j) + mat1(i, l) * mat2(l, j)
            l = l + 1
            Loop
        j = j + 1
        l = 0
        Loop
        i = i + 1
        j = 0
    Loop
      
'    print input and results to immediate window
'    For i = 0 To 5
'    Debug.Print mat1(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5) & vbTab * z(i, 0) & vbTab & z(i, 1) & vbTab & z(i, 2) & vbTab & z(i, 3) & vbTab & z(i, 4) & vbTab & z(i, 5)
'    Next i
'
   mult2matrix = z
    


    End Function
