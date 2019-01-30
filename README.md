![alt text](https://github.com/alexthomas96/Excel_Cell_Shifter/blob/master/excelicon.png "Image Source : https://www.flaticon.com/free-icon/excel_732220#term=excel&page=1&position=1")


# Excel_Cell_Shifter 
Code to shift excel cells and return cell name for a given offset. 
> eg: Shift cell "AB32" by offset 3 results in "AE32".

***

### Usage : 
* in_Cell_Name _(Input Paramter)_ : Name of the current cell as a string 
> eg : "AB32"
                                
* index_Offset _(Input Parameter)_ : Offset to shift as an integer (positive and negative)
> eg : 3
                                 
* out_Result _(Output Parameter)_ : The shifted result as a string
> eg : "AE32"

***

Returns error `INVALID_ENTRY` in case the value falls below column A. 
