---
title: Workbook.Open Event (Excel)
keywords: vbaxl10.chm503073
f1_keywords:
- vbaxl10.chm503073
ms.prod: excel
api_name:
- Excel.Workbook.Open
ms.assetid: 313adc5e-0319-4ca4-cf5d-791b7184dacf
ms.date: 06/08/2017
localization_priority: Priority
---


# Workbook.Open Event (Excel)

Occurs when the workbook is opened.


## Syntax

_expression_. `Open`

 _expression_ An expression that returns a [Workbook](./Excel.Workbook.md) object.


## Example

This example maximizes Microsoft Excel whenever the workbook is opened.


```vb
Private Sub Workbook_Open() 
 Application.WindowState = xlMaximized 
End Sub
```


## See also


[Workbook Object](Excel.Workbook.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]