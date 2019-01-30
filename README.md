
# Excel_Cell_Shifter ![alt text](https://github.com/alexthomas96/Excel_Cell_Shifter/blob/master/excelicon.png "Image Source : <https://camo.githubusercontent.com/b717aba80a45ddfd295ee23526382eeae2c0eb77/687474703a2f2f61313835322e70686f626f732e6170706c652e636f6d2f75732f7233302f507572706c65332f76342f36622f37622f62392f36623762623939392d316561322d393333322d653031612d6632383939313730633964392f6d7a6c2e626665766e73737a2e706e67>")

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
