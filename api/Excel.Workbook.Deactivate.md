---
title: Workbook.Deactivate event (Excel)
keywords: vbaxl10.chm503075
f1_keywords:
- vbaxl10.chm503075
ms.prod: excel
api_name:
- Excel.Workbook.Deactivate
ms.assetid: 6bd5411c-ac43-95cf-6755-49780ac765e9
ms.date: 05/29/2019
localization_priority: Normal
---


# Workbook.Deactivate event (Excel)

Occurs when the chart, worksheet, or workbook is deactivated.


## Syntax

_expression_.**Deactivate**

_expression_ A variable that represents a **[Workbook](Excel.Workbook.md)** object.


## Return value

**Nothing**


## Example

This example arranges all open windows when the workbook is deactivated.


```vb
Private Sub Workbook_Deactivate() 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]