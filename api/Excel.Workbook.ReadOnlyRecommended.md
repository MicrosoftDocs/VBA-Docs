---
title: Workbook.ReadOnlyRecommended property (Excel)
keywords: vbaxl10.chm199216
f1_keywords:
- vbaxl10.chm199216
ms.prod: excel
api_name:
- Excel.Workbook.ReadOnlyRecommended
ms.assetid: 3cae84e4-d5f0-f01c-64d9-ec586ffdf79c
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.ReadOnlyRecommended property (Excel)

**True** if the workbook was saved as read-only recommended. Read-only **Boolean**.


## Syntax

_expression_.**ReadOnlyRecommended**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Remarks

When you open a workbook that was saved as read-only recommended, Microsoft Excel displays a message recommending that you open the workbook as read-only.

Use the **[SaveAs](Excel.Workbook.SaveAs.md)** method to change this property.


## Example

This example displays a message if the active workbook is saved as read-only recommended.

```vb
If ActiveWorkbook.ReadOnlyRecommended = True Then 
 MsgBox "This workbook is saved as read-only recommended" 
End If
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]