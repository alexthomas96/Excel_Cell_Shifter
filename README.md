# Excel_Cell_Shifter
Code to shift excel cells and return cell name for a given offset. 
eg: Shift cell "AB32" by offset 3 results in "AE32".

### Usage : 
* in_Cell_Name (Input Paramter) : Name of the current cell as a string 
                                eg : "AB32"
                                
* index_Offset (Input Parameter) : Offset to shift as an integer (positive and negative)
                                 eg : 3
                                 
* out_Result (Output Parameter) : The shifted result as a string
                                eg : "AE32"

Returns error "INVALID_ENTRY" in case the value falls below column A. 
