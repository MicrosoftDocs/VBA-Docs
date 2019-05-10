---
title: Range.Areas property (Excel)
keywords: vbaxl10.chm144081
f1_keywords:
- vbaxl10.chm144081
ms.prod: excel
api_name:
- Excel.Range.Areas
ms.assetid: 31fc03b4-25b6-27ae-2350-b34c6c6ba255
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.Areas property (Excel)

Returns an **[Areas](Excel.Areas.md)** collection that represents all the ranges in a multiple-area selection. Read-only.


## Syntax

_expression_.**Areas**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

For a single selection, the **Areas** property returns a collection that contains one object—the original **Range** object itself. 

For a multiple-area selection, the **Areas** property returns a collection that contains one object for each selected area.


## Example

This example displays a message if the user tries to carry out a command when more than one area is selected. This example must be run from a worksheet.

```vb
If Selection.Areas.Count > 1 Then 
 MsgBox "Cannot do this to a multi-area selection." 
End If
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
