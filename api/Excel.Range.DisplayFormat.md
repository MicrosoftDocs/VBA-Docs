---
title: Range.DisplayFormat property (Excel)
keywords: vbaxl10.chm144251
f1_keywords:
- vbaxl10.chm144251
ms.prod: excel
api_name:
- Excel.Range.DisplayFormat
ms.assetid: c4e044e2-a04e-b655-2973-7e02897ca49d
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.DisplayFormat property (Excel)

Returns a **[DisplayFormat](Excel.DisplayFormat.md)** object that represents the display settings for the specified range. Read-only.


## Syntax

_expression_.**DisplayFormat**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Return value

DisplayFormat


## Remarks

Note that the **DisplayFormat** property does not work in user-defined functions. For example, on a worksheet function that returns the interior color of a cell, you use a line similar to: `Range(n).DisplayFormat.Interior.ColorIndex`. When the worksheet function executes, it returns a **#VALUE!** error.

In another example, you cannot use the **DisplayFormat** property in a worksheet function to return settings for a particular range. **DisplayFormat** will work in a function called from Visual Basic for Applications (VBA), however. For example, in the following function:

```vb
Function getColorIndex()
   getColorIndex = ActiveCell.DisplayFormat.Interior.ColorIndex
End Function
```

Calling the function from a worksheet as follows **=getColorIndex()** returns the **#VALUE!** error.

However, when the function is called from the Immediate pane in the Visual Basic Editor, it returns the color index for the active cell. To work around this issue, remove **DisplayFormat** from the code. The following function will work either from a worksheet or from VBA.

```vb
Function getColorIndex()
   getColorIndex = ActiveCell.Interior.ColorIndex
End Function
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
