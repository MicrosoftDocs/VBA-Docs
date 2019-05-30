---
title: Worksheet.Deactivate event (Excel)
keywords: vbaxl10.chm502077
f1_keywords:
- vbaxl10.chm502077
ms.prod: excel
api_name:
- Excel.Worksheet.Deactivate
ms.assetid: 3f66b86b-d0f0-bdc0-594c-3eb9faa44ff2
ms.date: 05/30/2019
localization_priority: Normal
---


# Worksheet.Deactivate event (Excel)

Occurs when the chart, worksheet, or workbook is deactivated.


## Syntax

_expression_.**Deactivate**

_expression_ A variable that represents a **[Worksheet](Excel.Worksheet.md)** object.


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
