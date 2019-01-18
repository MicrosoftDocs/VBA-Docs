---
title: Application.Quit method (Excel)
keywords: vbaxl10.chm133194
f1_keywords:
- vbaxl10.chm133194
ms.prod: excel
api_name:
- Excel.Application.Quit
ms.assetid: d01de494-95c7-6e3e-3049-f89b31aa9d0c
ms.date: 06/08/2017
localization_priority: Priority
---


# Application.Quit method (Excel)

Quits Microsoft Excel.


## Syntax

_expression_. `Quit`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Remarks

If unsaved workbooks are open when you use this method, Microsoft Excel displays a dialog box asking whether you want to save the changes. You can prevent this by saving all workbooks before using the  **Quit** method or by setting the **[DisplayAlerts](Excel.Application.DisplayAlerts.md)** property to **False**. When this property is **False** , Microsoft Excel doesn't display the dialog box when you quit with unsaved workbooks; it quits without saving them.

If you set the  **[Saved](Excel.Workbook.Saved.md)** property for a workbook to **True** without saving the workbook to the disk, Microsoft Excel will quit without asking you to save the workbook.


## Example

This example saves all open workbooks and then quits Microsoft Excel.


```vb
For Each w In Application.Workbooks 
 w.Save 
Next w 
Application.Quit
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]