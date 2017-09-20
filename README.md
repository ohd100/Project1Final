Matrix Functions
Below is the name of each function or macro, associated data types, and what it does:

Function addmatrix(mat1 As Variant, mat2 As Variant) As Variant  
'    This function adds the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are the same;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

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


Function subtmatrix(mat1 As Variant, mat2 As Variant) As Variant
'    This function subtracts the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are the same;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

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

Function mult2matrix(mat1 As Variant, mat2 As Variant) As Variant
'    This function multiplies the two inputted matrices (mat1 and mat2)
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are compatible for multiplication;
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

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

Function divmatrix(mat1 As Variant, mat2 As Variant) As Variant
'    This function divides the two inputted matrices (mat1 and mat2) (multiplies mat1 by the inverse of mat2)
'    Note that both inputs have to be in the variant data format
'    Utilized the "MInverse" VBA built-in VBA function
'    Utilized "multmatfordiv" function instead of "mult2matrix" function due to debugging errors I was getting related to how the inverse function creates a solution matrix (labels it starting from 1 instead of 0)
'    Added functionality where function checks if dimensions of input matrices are compatible for division
'       -> mat2 must be a square matrix and #mat1columns must = #mat2rows
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

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

Function multmatfordiv(mat1 As Variant, mat2 As Variant) As Variant
'    This function multiplies the two inputted matrices (mat1 and mat2)
'    Seperate from other matrix multiplication function due to debugging issues I was having early in the coding process due to how the MInverse function creates the solution matrix
'    Note that both inputs have to be in the variant data format
'    Added functionality where function checks if dimensions of input matrices are compatible for multiplication (number of columns in mat1 has to equal the number of rows in mat2);
'       -> If dimensions are different, function will end and throw a messagebox saying "Matrix size mismatch error"

Function WriteArrayToImmediateWindow(arrSubA As Variant)
'   borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'   Made some modifications to get it to display all arrays properly

Function WriteArraydiv(arrSubA As Variant)
'   borrowed this function from https://stackoverflow.com/questions/14274949/how-to-print-two-dimensional-array-in-immediate-window-in-vba
'   Made some modifications to get it to display all arrays properly
'   Seperate from other array printing function due to debugging issues I was having early in the coding process with the division function
