'in_Cell_Name (Input Paramter) : Name Of the current cell As a String eg : "AB32"
' // in code inArg = in_Cell_Name

'index_Offset (Input Parameter) : Offset To shift As an Integer eg : 3
' // in code inOffset = index_Offset

'out_Result (Output Parameter) : The shifted result As a String eg : "AE32"
' // in code outCell = out_Result

Dim cellArray() As Char 'Define a cell array 
cellArray = System.Text.RegularExpressions.Regex.Replace(inArg,"[^/A-Z,a-z]","").ToCharArray 'Replace all non-alphabetic characters with null
outArg = Convert.ToInt32((1 - (26 ^ cellArray.Length)) / (1 - 26)) 'Calculate number of sets to traverse the excel structure. eg : for input of cell "ABC23", skip though sections A-Z, AA-ZZ, AAA-AAZ, ABA and ABB
Dim arrayItem As Char
Dim loopVar As Integer
loopVar = 0
'generate column number of given cell
For Each arrayItem In cellArray 'loop through each array element
 If loopVar = (cellArray.Length - 1) Then
          outArg = outArg + (Asc(arrayItem) - 65) 'calculate spacing / offset by difference of ASCII values for last item
      Else
          outArg = outArg + ((Asc(arrayItem) - 65) * Convert.ToInt32(26 ^ (cellArray.Length - loopVar - 1))) 'calculate spacing / offset by difference of ASCII values for other items by exponentiation to base 26
      End If
   loopVar = loopVar + 1
Next

Dim Column_Number As Integer
Dim Column_Name As String
Column_Name = ""
Dim modulo As Integer
Column_Number = outArg + inOffset 'shifted value by adding offset with column number 
While Column_Number > 0
    modulo = (Column_Number - 1) Mod 26
 Column_Name = Chr(65 + modulo).ToString() & Column_Name 'obtain column name
 Column_Number = Convert.ToInt32((Column_Number - modulo) / 26) 'obtain final column number
End While
  ' Your functions / subs goes here
outCell = Column_Name & System.Text.RegularExpressions.Regex.Replace(inArg,"[^/0-9]","").ToString
If (Column_Number < 0 Or System.Text.RegularExpressions.Regex.Replace(outCell,"[0-9]","").ToString = "" )
 outCell = "INVALID_ENTRY" 'error if value falls below column A (first column)
End If
