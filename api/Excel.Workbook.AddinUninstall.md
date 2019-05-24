---
title: Workbook.AddinUninstall event (Excel)
keywords: vbaxl10.chm503081
f1_keywords:
- vbaxl10.chm503081
ms.prod: excel
api_name:
- Excel.Workbook.AddinUninstall
ms.assetid: e35ba67b-3e04-d950-2f8b-141e478ddb67
ms.date: 05/25/2019
localization_priority: Normal
---


# Workbook.AddinUninstall event (Excel)

Occurs when the workbook is uninstalled as an add-in.


## Syntax

_expression_.**AddinUninstall**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Nothing**


## Remarks

The add-in doesn't automatically close when it's uninstalled.


## Example

This example minimizes Microsoft Excel when the workbook is uninstalled as an add-in.

```vb
Private Sub Workbook_AddinUninstall() 
 Application.WindowState = xlMinimized 
End Sub
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]