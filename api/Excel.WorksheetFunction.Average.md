---
title: WorksheetFunction.Average Method (Excel)
keywords: vbaxl10.chm137078
f1_keywords:
- vbaxl10.chm137078
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Average
ms.assetid: 9d6b697d-f7e0-6e81-a4a4-39fafafb879f
ms.date: 06/08/2017
---


# WorksheetFunction.Average Method (Excel)

Returns the average (arithmetic mean) of the arguments.


## Syntax

 _expression_. `Average`( `_Arg1_` , `_Arg2_` , `_Arg3_` , `_Arg4_` , `_Arg5_` , `_Arg6_` , `_Arg7_` , `_Arg8_` , `_Arg9_` , `_Arg10_` , `_Arg11_` , `_Arg12_` , `_Arg13_` , `_Arg14_` , `_Arg15_` , `_Arg16_` , `_Arg17_` , `_Arg18_` , `_Arg19_` , `_Arg20_` , `_Arg21_` , `_Arg22_` , `_Arg23_` , `_Arg24_` , `_Arg25_` , `_Arg26_` , `_Arg27_` , `_Arg28_` , `_Arg29_` , `_Arg30_` )

 _expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


### Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1 - Arg30_|Required| **Variant**|1 to 30 numeric arguments for which you want the average.|

### Return value

Double


## Remarks


-  Arguments can either be numbers or names, arrays, or references that contain numbers.
    
- Logical values and text representations of numbers that you type directly into the list of arguments are counted.
    
- If an array or reference argument contains text, logical values, or empty cells, those values are ignored; however, cells with the value zero are included.
    
- Arguments that are error values or text that cannot be translated into numbers cause errors.
    
- If you want to include logical values and text representations of numbers in a reference as part of the calculation, use the AVERAGEA function.
    

 **Note**  The Average method measures central tendency, which is the location of the center of a group of numbers in a statistical distribution. The three most common measures of central tendency are:


-  **Average** which is the arithmetic mean, and is calculated by adding a group of numbers and then dividing by the count of those numbers. For example, the average of 2, 3, 3, 5, 7, and 10 is 30 divided by 6, which is 5.
    
-  **Median** which is the middle number of a group of numbers; that is, half the numbers have values that are greater than the median, and half the numbers have values that are less than the median. For example, the median of 2, 3, 3, 5, 7, and 10 is 4.
    
-  **Mode** which is the most frequently occurring number in a group of numbers. For example, the mode of 2, 3, 3, 5, 7, and 10 is 3.
    
For a symmetrical distribution of a group of numbers, these three measures of central tendency are all the same. For a skewed distribution of a group of numbers, they can be different.

 **Tip** When averaging cells, keep in mind the difference between empty cells and those containing the value zero, especially if you have cleared the **Zero values** check box on the **View** tab (**Options** command, **Tools** menu). Empty cells are not counted, but zero values are.


## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

