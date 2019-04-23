---
title: CubeField.TreeviewControl property (Excel)
keywords: vbaxl10.chm668079
f1_keywords:
- vbaxl10.chm668079
ms.prod: excel
api_name:
- Excel.CubeField.TreeviewControl
ms.assetid: 54f44b41-cde8-aa06-af98-c7d79fc85c12
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.TreeviewControl property (Excel)

Returns the **[TreeviewControl](Excel.TreeviewControl.md)** object of the **CubeField** object, representing the cube manipulation control of an OLAP-based PivotTable report. Read-only.


## Syntax

_expression_.**TreeviewControl**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

This property is available only when the control is visible.


## Example

This example sets the first cube field control to "drilled" for the states of California and Maryland in the second PivotTable report on the active worksheet.

```vb
ActiveSheet.PivotTables("PivotTable2") _ 
 .CubeFields(1).TreeviewControl.Drilled = _ 
 Array(Array("", ""), _ 
 Array("[state].[states].[CA]", _ 
 "[state].[states].[MD]"))
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]