---
title: Chart.Deactivate Event (Excel)
keywords: vbaxl10.chm500074
f1_keywords:
- vbaxl10.chm500074
ms.prod: excel
api_name:
- Excel.Chart.Deactivate
ms.assetid: b843b64a-ad20-d160-1abb-88317114b44c
ms.date: 06/08/2017
localization_priority: Normal
---


# Chart.Deactivate Event (Excel)

Occurs when the chart, worksheet, or workbook is deactivated.


## Syntax

_expression_. `Deactivate`

_expression_ A variable that represents a [Chart](Excel.Chart-graph-object.md) object.


## Example

This example arranges all open windows when the workbook is deactivated.


```vb
Private Sub Workbook_Deactivate() 
 Application.Windows.Arrange xlArrangeStyleTiled 
End Sub
```


## See also


[Chart Object](Excel.Chart(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]