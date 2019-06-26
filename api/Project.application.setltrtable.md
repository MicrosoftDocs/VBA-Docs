---
title: Application.SetLTRTable method (Project)
keywords: vbapj.chm1520
f1_keywords:
- vbapj.chm1520
ms.prod: project-server
ms.assetid: 33aee9ba-da55-c83c-a1cf-27b5751c3fdf
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.SetLTRTable method (Project)
Sets column order from left to right, for a selected table in a report.

## Syntax

_expression_. `SetLTRTable`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Return value

 **Boolean**

 **True** if the column order is set from left to right; otherwise, **False**.


## Remarks

The  **SetLTRTable** method can be used to change the table columns from right-to-left order for languages such as Arabic, to left-to-right for languages such as English, German, and French.

If a report is not active, the  **SetLTRTable** method displays a dialog box with run-time error 1100, "The method is not available in this situation."


## See also


[Application Object](Project.Application.md)



[SetRTLTable](Project.application.setrtltable.md)
[ReportTable Object](Project.reporttable.md)
[Shape.Table Property](Project.shape.table.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]