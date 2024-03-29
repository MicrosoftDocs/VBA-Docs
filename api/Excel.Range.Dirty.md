---
title: Range.Dirty method (Excel)
keywords: vbaxl10.chm144234
f1_keywords:
- vbaxl10.chm144234
api_name:
- Excel.Range.Dirty
ms.assetid: c3f177ef-19b9-07e7-a42f-978874528207
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.Dirty method (Excel)

Designates a range to be recalculated when the next recalculation occurs.


## Syntax

_expression_.**Dirty**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The **[Calculate](Excel.Range.Calculate.md)** method forces the specified range to be recalculated for cells that Microsoft Excel understands as needing recalculation.

If the application is in manual calculation mode, using the **Dirty** method instructs Excel to identify the specified cell to be recalculated. If the application is in automatic calculation mode, using the **Dirty** method instructs Excel to perform a recalculation.


## Example

In this example, Excel enters a formula in cell A3, saves the changes, and then recalculates cell A3.

```vb
Sub UseDirtyMethod() 
 
 MsgBox "Two values and a formula will be entered." 
 Range("A1").Value = 1 
 Range("A2").Value = 2 
 Range("A3").Formula = "=A1+A2" 
 
 ' Save the changes made to the worksheet. 
 Application.DisplayAlerts = False 
 Application.Save 
 MsgBox "Changes saved." 
 
 ' Force a recalculation of range A3. 
 Application.Range("A3").Dirty 
 MsgBox "Try to close the file without saving and a dialog box will appear." 
 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
