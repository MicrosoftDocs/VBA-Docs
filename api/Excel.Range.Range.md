---
title: Range.Range property (Excel)
keywords: vbaxl10.chm144184
f1_keywords:
- vbaxl10.chm144184
ms.prod: excel
api_name:
- Excel.Range.Range
ms.assetid: 7edbda7c-98d9-143d-7b5e-bcfb7f237818
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.Range property (Excel)

Returns a **Range** object that represents a cell or a range of cells.


## Syntax

_expression_.**Range** (_Cell1_, _Cell2_)

_expression_ A variable that represents a **[Range](Excel.Range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Cell1_|Required| **Variant**|The name of the range. This must be an A1-style reference in the language of the macro. It can include the range operator (a colon), the intersection operator (a space), or the union operator (a comma). It can also include dollar signs, but they're ignored.<br/><br/>You can use a local defined name in any part of the range. If you use a name, the name is assumed to be in the language of the macro.|
| _Cell2_|Optional| **Variant**|The cell in the upper-left and lower-right corner of the range. Can be a **Range** object that contains a single cell, an entire column, or entire row, or it can be a string that names a single cell in the language of the macro.|

## Remarks

When used without an object qualifier, this property is a shortcut for **[ActiveSheet.Range](Excel.Worksheet.Range.md)** (it returns a range from the active sheet; if the active sheet isn't a worksheet, the property fails).

When applied to a **Range** object, the property is relative to the **Range** object. For example, if the selection is cell C3, `Selection.Range("B1")` returns cell D3 because it's relative to the **Range** object returned by the **Selection** property. On the other hand, the code `ActiveSheet.Range("B1")` always returns cell B1.


## Example

This example sets the value of the top-left cell of the range B2:C4 on Sheet1 of the active workbook, i.e. that of the cell B2, to 3.14159.

```vb
With Worksheets("Sheet1").Range("B2:C4")
   .Range("A1").Value = 3.14159
End With
```

<br/>

This example loops on the the four cells in the top-left corner of the range B2:Z22 on Sheet1 of the active workbook. If one of the cells has a value less than 0.001, the code replaces that value with 0 (zero).

```vb
Public Sub TruncateSmallValues()
   Dim exampleRange As Excel.Range
   Set exampleRange = Worksheets("Sheet1").Range("B2:Z22") 

   Dim cell As Excel.Range
   For Each cell in exampleRange.Range("A1:B2") 
      If cell.Value < .001 Then 
         cell.Value = 0 
      End If 
   Next cell
End Sub
```

<br/>

This example sets the font style in cells B2:D6 on Sheet1 of the active workbook to italic. The example uses Syntax 2 of the **Range** property.

```vb
With Worksheets("Sheet1").Range("B2:Z22")
   .Range(.Cells(1, 1), .Cells(5, 3)).Font.Italic = True 
End With
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
