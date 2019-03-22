---
title: Screen.ActiveReport property (Access)
keywords: vbaac10.chm12491
f1_keywords:
- vbaac10.chm12491
ms.prod: access
api_name:
- Access.Screen.ActiveReport
ms.assetid: efcf6bfd-2749-5b5c-d7ca-a26168bfcb65
ms.date: 03/23/2019
localization_priority: Normal
---


# Screen.ActiveReport property (Access)

You can use the **ActiveReport** property together with the **Screen** object to identify or refer to the report that has the focus. Read-only **Report** object.


## Syntax

_expression_.**ActiveReport**

_expression_ A variable that represents a **[Screen](Access.Screen.md)** object.


## Remarks

This property setting contains a reference to the **[Report](Access.Report.md)** object that has the focus at run time.

You can use the **ActiveReport** property to refer to an active report together with one of its properties or methods. The following example displays the **Name** property setting of the active report.

```vb
Dim rptCurrentReport As Report 
Set rptCurrentReport = Screen.ActiveReport 
MsgBox "Current report is " & rptCurrentReport.Name
```

If no report has the focus when you use the **ActiveReport** property, an error occurs.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]