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

**DisplayFormat** is affected by conditional formatting as shown in the code below. It adds conditional formatting to Cell A1 on the ActiveSheet. This formatting bolds the cell, changes the interior color to red and adds a checker pattern.
```vb
Public Sub DemonstrateConditionalFormattingAffectsDisplayFormat()
    Dim inputArea As Range
    Set inputArea = ActiveSheet.Range("A1")
    
    Dim addedFormatCondition As FormatCondition
    Set addedFormatCondition = inputArea.FormatConditions.Add(xlExpression, Formula1:="=true")
    addedFormatCondition.Font.Bold = True
    addedFormatCondition.Interior.Color = XlRgbColor.rgbRed
    addedFormatCondition.Interior.Pattern = XlPattern.xlPatternChecker
    
    Debug.Print inputArea.Font.Bold 'False
    Debug.Print inputArea.Interior.Color 'XlRgbColor.rgbWhite
    Debug.Print inputArea.Interior.Pattern 'XlPattern.xlPatternNone
    
    Debug.Print inputArea.DisplayFormat.Font.Bold 'True
    Debug.Print inputArea.DisplayFormat.Interior.Color 'XlRgbColor.rgbRed
    Debug.Print inputArea.DisplayFormat.Interior.Pattern 'XlPattern.xlPatternChecker
End Sub
```

Note that the **DisplayFormat** property does not work in User Defined Functions (UDF). For example, on a worksheet function that returns the interior color of a cell, you use a line similar to: `Range(n).DisplayFormat.Interior.ColorIndex`. When the worksheet function executes, it returns a **#VALUE!** error.

In another example, you cannot use the **DisplayFormat** property in a worksheet function to return settings for a particular range. **DisplayFormat** will work in a function called from Visual Basic for Applications (VBA), however. For example, in the following UDF:

```vb
Function getDisplayedColorIndex()
   getColorIndex = ActiveCell.DisplayFormat.Interior.ColorIndex
End Function
```

Calling the function from a worksheet as follows **=getDisplayedColorIndex()** returns the **#VALUE!** error. As such, if conditional formatting is applied to a range, there is no way to return that value with a UDF. If conditional formatting has been applied, obtain the color index for the active cell by calling the Immediate pane in the Visual Basic Editor.

If no conditional formatting is applied use the function below to returns the color index for the active cell. The following function will work either from a worksheet or from VBA.

```vb
Function getAppliedColorIndex()
   getColorIndex = ActiveCell.Interior.ColorIndex
End Function
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
