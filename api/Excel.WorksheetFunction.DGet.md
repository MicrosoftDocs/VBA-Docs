---
title: WorksheetFunction.DGet method (Excel)
keywords: vbaxl10.chm137170
f1_keywords:
- vbaxl10.chm137170
ms.prod: excel
api_name:
- Excel.WorksheetFunction.DGet
ms.assetid: 71c12527-19a6-7fb7-b1c1-f2b5478c14b9
ms.date: 06/08/2017
localization_priority: Normal
---


# WorksheetFunction.DGet method (Excel)

Extracts a single value from a column of a list or database that matches conditions that you specify.


## Syntax

_expression_. `DGet`( `_Arg1_` , `_Arg2_` , `_Arg3_` )

_expression_ A variable that represents a [WorksheetFunction](./Excel.WorksheetFunction.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|Database - the range of cells that makes up the list or database. A database is a list of related data in which rows of related information are records, and columns of data are fields. The first row of the list contains labels for each column.|
| _Arg2_|Required| **Variant**|Field - indicates which column is used in the function. Enter the column label enclosed between double quotation marks, such as "Age" or "Yield," or a number (without quotation marks) that represents the position of the column within the list: 1 for the first column, 2 for the second column, and so on.|
| _Arg3_|Required| **Variant**|Criteria - the range of cells that contains the conditions that you specify. You can use any range for the criteria argument, as long as it includes at least one column label and at least one cell below the column label in which you specify a condition for the column.|

## Return value

Variant


## Remarks




- Because the equal sign is used to indicate a formula when you type text or a value in a cell, Microsoft Excel evaluates what you type; however, this may cause unexpected filter results. To indicate an equality comparison operator for either text or a value, type the criteria as a string expression in the appropriate cell in the criteria range: **=''=**_entry_**''**Where  _entry_ is the text or value you want to find. For example:
    

|**What you type in the cell**|**What Excel evaluates and displays**|
|:-----|:-----|
|="=Davolio"|=Davolio|
|="=3000"|=3000|

- When filtering text data, Excel does not distinguish between uppercase and lowercase characters. However, you can use a formula to perform a case-sensitive search.
    

## See also


[WorksheetFunction Object](Excel.WorksheetFunction.md)

