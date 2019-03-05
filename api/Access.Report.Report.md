---
title: Report.Report property (Access)
keywords: vbaac10.chm13791
f1_keywords:
- vbaac10.chm13791
ms.prod: access
api_name:
- Access.Report.Report
ms.assetid: 0cacc875-2083-159a-423f-757ab19e5839
ms.date: 03/06/2019
localization_priority: Normal
---


# Report.Report property (Access)

You can use the **Report** property to refer to a report or to refer to the report associated with a subreport control. Read-only **Report**.


## Syntax

_expression_.**Report**

_expression_ A variable that represents a **[Report](Access.Report.md)** object.


## Remarks

This property is typically used to refer to the report contained in a subreport control.

When you use the **[Reports](Access.Reports.md)** collection, you must specify the name of the report.


## Example

The following example uses the **Report** property to refer to a control on a subreport.

```vb
Dim curTotalSales As Currency 
 
curTotalSales = Reports!Sales!Employees.Report!TotalSales
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]