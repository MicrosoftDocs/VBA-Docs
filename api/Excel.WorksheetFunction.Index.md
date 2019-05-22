---
title: WorksheetFunction.Index method (Excel)
keywords: vbaxl10.chm137090
f1_keywords:
- vbaxl10.chm137090
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Index
ms.assetid: 4656985a-2864-93ed-31c7-e7a551d68e96
ms.date: 05/23/2019
localization_priority: Normal
---


# WorksheetFunction.Index method (Excel)

Returns a value or the reference to a value from within a table or range. There are two forms of the **Index** function: the array form and the reference form.


## Syntax

_expression_.**Index** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Array or Reference - a range of cells or an array constant. For references, it is the reference to one or more cell ranges.|
| _Arg2_|Required| **Double**|Row_num - selects the row in array from which to return a value. If row_num is omitted, column_num is required. For references, the number of the row in reference from which to return a reference.|
| _Arg3_|Optional| **Variant**|Column_num - selects the column in array from which to return a value. If column_num is omitted, row_num is required. For reference, the number of the column in reference from which to return a reference.|
| _Arg4_|Optional| **Variant**|Area_num - only used when returning references. Selects a range in reference from which to return the intersection of row_num and column_num. The first area selected or entered is numbered 1, the second is 2, and so on. If area_num is omitted, **Index** uses area 1.|

## Return value

**Variant**


## Remarks

### Array form

Returns the value of an element in a table or an array, selected by the row and column number indexes.

Use the array form if the first argument to **Index** is an array constant.

If both the row_num and column_num arguments are used, **Index** returns the value in the cell at the intersection of row_num and column_num.
    
If you set row_num or column_num to 0 (zero), **Index** returns the array of values for the entire column or row, respectively. To use values returned as an array, enter the **Index** function as an array formula in a horizontal range of cells for a row, and in a vertical range of cells for a column. To enter an array formula, press Ctrl+Shift+Enter.
    
Row_num and column_num must point to a cell within array; otherwise, **Index** returns the #REF! error value.
    

### Reference form

Returns the reference of the cell at the intersection of a particular row and column. If the reference is made up of nonadjacent selections, you can pick the selection to look in. If each area in reference contains only one row or column, the row_num or column_num argument, respectively, is optional. For example, for a single row reference, use INDEX(reference,column_num). 

After reference and area_num have selected a particular range, row_num and column_num select a particular cell: row_num 1 is the first row in the range, column_num 1 is the first column, and so on. The reference returned by **Index** is the intersection of row_num and column_num.
    
If you set row_num or column_num to 0 (zero), **Index** returns the reference for the entire column or row, respectively. 
    
Row_num, column_num, and area_num must point to a cell within reference; otherwise, **Index** returns the #REF! error value. If row_num and column_num are omitted, **Index** returns the area in reference specified by area_num.
    
The result of the **Index** function is a reference and is interpreted as such by other formulas. Depending on the formula, the return value of **Index** may be used as a reference or as a value. For example, the formula `CELL("width",INDEX(A1:B2,1,2))` is equivalent to `CELL("width",B1)`. The CELL function uses the return value of **Index** as a cell reference. On the other hand, a formula such as `2*INDEX(A1:B2,1,2)` translates the return value of **Index** into the number in cell B1.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
