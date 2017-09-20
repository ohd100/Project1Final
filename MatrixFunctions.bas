Attribute VB_Name = "Module1"
    Dim a As Integer, b As Integer, c As Integer, d As Integer, i As Integer, j As Integer
    Dim matsol As Integer

Function addmatrix(mat1 As Variant, mat2 As Variant) As Variant
    
'    This function adds the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are the same;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"
    
    
    Dim z() As Long
    
    
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   mat2row = UBound(mat2, 1)
   
    If mat1row = mat2row Then
        If mat1col = mat1col Then
   
           ReDim z(mat1row, mat1col)
           
           a = mat1row
           b = mat1col
           d = mat2col
         
           
           i = 0
           j = 0
        
            
            
            Do While i <= a                                 ''solution matrix
                Do While j <= b
                    z(i, j) = mat1(i, j) + mat2(i, j)
                    j = j + 1
                Loop
                i = i + 1
                j = 0
            Loop
              
            addmatrix = z
    
        Else: MsgBox "Matrix size mismatch error"
        End If
    Else: MsgBox "Matrix size mismatch error"
    End If

    End Function

Sub runaddmatrix()
    
'    This sub (macro) creates two matrices (mat1 and mat2)
'       Dimensions of both matrices are specified as follows below:
'           mat1 rows is variable mat1row+1 (set to 5 currently so # of rows is 6)
'           mat2 rows is variable mat2row+1 (set to 5 currently so # of rows is 6)
'           mat1 columns is variable m1col+1 (set to 5 currently so # of columns is 6)
'           mat1 columns is variable mat2col+1 (set to 5 currently so # of columns is 6)
'
'    Only matrices of the same size can be added or subtracted, so if the sizes are different, a message box will popup saying "Matrix size mismatch error"
'
'   Then, each position within both matrices is initialized with random values (integers from 0 to 100)
'
'   Then, these matrices are inputted to the addmatrix function defined in this module.
'
'   Then, the input matrices and the solution matrix are printed to the immediate window to verify results

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
        
Debug.Print WriteArrayToImmediateWindow(mat1)       'Prints mat1 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(mat2)       'Prints mat2 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(z)          'Prints solution matrix to "Immediate Window"

End Sub

Function subtmatrix(mat1 As Variant, mat2 As Variant) As Variant
   
'    This function subtracts the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are the same;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"
    
   
   Dim z() As Long
    
    
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   mat2row = UBound(mat2, 1)
   
   
   If mat1row = mat2row Then
        If mat1col = mat1col Then
   
               ReDim z(mat1row, mat1col)
               
               a = mat1row
               b = mat1col
               d = mat2col
              
               
               i = 0
               j = 0
            
                
                
                Do While i <= a                                 ''solution matrix
                    Do While j <= b
                        z(i, j) = mat1(i, j) - mat2(i, j)
                        j = j + 1
                    Loop
                    i = i + 1
                    j = 0
                Loop
                  
            
                subtmatrix = z
    
        Else: MsgBox "Matrix size mismatch error"
        End If
    Else: MsgBox "Matrix size mismatch error"
    End If
    
    End Function

Sub runsubtmatrix()
    
'    This sub (macro) creates two matrices (mat1 and mat2)
'       Dimensions of both matrices are specified as follows below:
'           mat1 rows is variable mat1row+1 (set to 5 currently so # of rows is 6)
'           mat2 rows is variable mat2row+1 (set to 5 currently so # of rows is 6)
'           mat1 columns is variable m1col+1 (set to 5 currently so # of columns is 6)
'           mat1 columns is variable mat2col+1 (set to 5 currently so # of columns is 6)
'
'    Only matrices of the same size can be added or subtracted, so if the sizes are different, a message box will popup saying "Matrix size mismatch error"
'
'   Then, each position within both matrices is initialized with random values (integers from 0 to 100)
'
'   Then, these matrices are inputted to the subtmatrix function defined in this module.
'
'   Then, the input matrices and the solution matrix are printed to the immediate window to verify results
    
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
   
   z = subtmatrix(mat1, mat2)                       'solution matrix
        
Debug.Print WriteArrayToImmediateWindow(mat1)       'Prints mat1 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(mat2)       'Prints mat2 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(z)          'Prints solution matrix to "Immediate Window"

End Sub


Function mult2matrix(mat1 As Variant, mat2 As Variant) As Variant

'    This function multiplies the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are compatible for multiplication;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"


   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   mat2row = UBound(mat2, 1)
   
   a = mat1row
   d = mat2col
   l = mat1col
   
   i = 0
   j = 0
   l = 0
   
   
   If mat1col = mat2row Then
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
      

            mult2matrix = z
    
    Else: MsgBox "Matrix size mismatch error"
    End If

    End Function
    
    Sub runmultmatrix()
    
'    This sub (macro) creates two matrices (mat1 and mat2)
'       Dimensions of both matrices are specified as follows below:
'           mat1 rows is variable mat1row+1 (set to 1 currently so # of rows is 2)
'           mat2 rows is variable mat2row+1 (set to 1 currently so # of rows is 2)
'           mat1 columns is variable m1col+1 (set to 2 currently so # of columns is 4)
'           mat1 columns is variable mat2col+1 (set to 2 currently so # of columns is 4)
'
'    Only matrices where #mat1columns = #mat2rows can be multiplied"
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

'   Then, each position within both matrices is initialized with random values (integers from 0 to 100)
'
'   Then, these matrices are inputted to the mult2matrix function defined in this module.
'
'   Then, the input matrices and the solution matrix are printed to the immediate window to verify results
      
    
    
    mat1row = 1
    m1col = 3
    mat2row = 3
    mat2col = 1
 
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
   
   z = mult2matrix(mat1, mat2)                      'solution matrix

Debug.Print WriteArrayToImmediateWindow(mat1)       'Prints mat1 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(mat2)       'Prints mat2 to "Immediate Window"
Debug.Print WriteArrayToImmediateWindow(z)          'Prints solution matrix to "Immediate Window"
    
End Sub

Function divmatrix(mat1 As Variant, mat2 As Variant) As Variant
   
'    This function divides the two inputted matrices (mat1 and mat2) (multiplies mat1 by the inverse of mat2)
'    Note that both inputs have to be in the variant data format
'    Utilized the "MInverse" VBA built-in VBA function
'    Utilized "multmatfordiv" function instead of "mult2matrix" function due to debugging errors I was getting related to how the inverse function creates a solution matrix (labels it starting from 1 instead of 0)
'    Added functionality where function checks if dimensions of input matrices are compatible for division
'       -> mat2 must be a square matrix and #mat1columns must = #mat2rows
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   mat2row = UBound(mat2, 1)

   i = 0
   j = 0

   If mat1col = mat2row Then
        If mat2row = mat2col Then
        
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
              
        
           divmatrix = multmatfordiv(mat1, m2invexp)
    
        Else: MsgBox "Matrix size mismatch error"
        End If
    Else: MsgBox "Matrix size mismatch error"
    End If

    End Function
    
Sub rundivmatrix()
    
'    This sub (macro) creates two matrices (mat1 and mat2)
'       Dimensions of both matrices are specified as follows below:
'           mat1 rows is variable mat1row+1 (set to 4 currently so # of rows is 5)
'           mat2 rows is variable mat2row+1 (set to 4 currently so # of rows is 5)
'           mat1 columns is variable m1col+1 (set to 4 currently so # of columns is 5)
'           mat1 columns is variable mat2col+1 (set to 4 currently so # of columns is 5)
'
'    Only matrices where #mat1columns = #mat2rows can be multiplied"
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

'   Then, each position within both matrices is initialized with random values (integers from 0 to 100)
'
'   Then, these matrices are inputted to the divmatrix function defined in this module.
'
'   Then, the input matrices and the solution matrix are printed to the immediate window to verify results
    
    mat1row = 4
    m1col = 4
    mat2row = 4
    mat2col = 4
 
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

    If c = d Then                                   ''initializing mat2 matrix with random values
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

  ReDim z(a, d)
   
   z = divmatrix(mat1, mat2)                        'solution matrix


Debug.Print WriteArraydiv(mat1)       'Prints mat1 to "Immediate Window"
Debug.Print WriteArraydiv(mat2)       'Prints mat2 to "Immediate Window"
Debug.Print WriteArraydiv(z)          'Prints solution matrix to "Immediate Window"
    
End Sub


Function multmatfordiv(mat1 As Variant, mat2 As Variant) As Variant

'    This function multiplies the two inputted matrices (mat1 and mat2)
'    Seperate from other matrix multiplication function due to debugging issues I was having early in the coding process due to how the MInverse function creates the solution matrix
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are compatible for multiplication (number of columns in mat1 has to equal the number of rows in mat2);
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"
   
   mat1row = UBound(mat1, 1)
   mat1col = UBound(mat1, 2)
   mat2col = UBound(mat2, 2)
   mat2row = UBound(mat2, 1)
   
   a = mat1row
   d = mat2col
   l = mat1col
   
   i = 0
   j = 0
   l = 0
   
    If mat1col = mat2row Then
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

        multmatfordiv = z
    Else: MsgBox "Matrix size mismatch error"
    End If

    End Function

'   borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'   made some modifications to get it to display all arrays properly

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
 
    
'   borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'   made some modifications to get it to display all arrays properly
'   Seperate from other array printing function due to debugging issues I was having early in the coding process with the division function

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
