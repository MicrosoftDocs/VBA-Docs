---
title: Range.Value property (Excel)
keywords: vbaxl10.chm144216
f1_keywords:
- vbaxl10.chm144216
ms.prod: excel
api_name:
- Excel.Range.Value
ms.assetid: 23f28b24-430a-6ea4-4895-0afff8dff218
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Value property (Excel)

Returns or sets a **Variant** value that represents the value of the specified range.


## Syntax

_expression_.**Value** (_RangeValueDataType_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _RangeValueDataType_|Optional| **Variant**|The range value data type. Can be an **[XlRangeValueDataType](Excel.XlRangeValueDataType.md)** constant.|

## Remarks

When setting a range of cells with the contents of an XML spreadsheet file, only values of the first sheet in the workbook are used. You cannot set or get a discontiguous range of cells in the XML spreadsheet format.

The default member of **Range** forwards calls without parameters to **Value**. Thus, `someRange = someOtherRange` is equivalent to `someRange.Value = someOtherRange.Value`.

For ranges whose first area contains more than one cell, **Value** returns a **Variant** containing a 2-dimensional array of the values in the individual cells of the first range.

Assigning a 2-dim array to the the **Value** property will copy the values to the range in one operation. If the target range is larger than the array, the remaining cells will receive an error value.

Assigning an array to a multi-area range is not properly supported and should be avoided.


## Example

This example sets the value of cell A1 on Sheet1 of the active workbook to 3.14159.

```vb
Worksheets("Sheet1").Range("A1").Value = 3.14159
```

<br/>

This example loops on cells A1:D10 on Sheet1 of the active workbook. If one of the cells has a value of less than 0.001, the code replaces the value with 0 (zero).

```vb
For Each cell in Worksheets("Sheet1").Range("A1:D10") 
   If cell.Value < .001 Then 
      cell.Value = 0 
   End If 
Next cell
```

<br/>

This example loops over the values in the range A1:CC5000 on Sheet1. If one of the values is less than 0.001, the code replaces the value with 0 (zero). Finally it copies the values to the original range.

```vb
Public Sub TruncateSmallValuesInDataArea()
   Dim dataArea As Excel.Range
   Set dataArea = ThisworkBook.Worksheets("Sheet1").Range("A1:CC5000")
   
   Dim valuesArray() As Variant
   valuesArray = dataArea.Value
   
   Dim rowIndex As Long
   Dim columnIndex As Long
   For rowIndex = LBound(valuesArray, 1) To UBound(valuesArray, 1)
      For columnIndex = LBound(valuesArray, 2) To UBound(valuesArray, 2)
	     If valuesArray(rowIndex, columnIndex) < 0.001 Then
		    valuesArray(rowIndex, columnIndex) = 0
		 End If 
	  Next
   Next
   
   dataArea.Value = valuesArray
End Sub
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
