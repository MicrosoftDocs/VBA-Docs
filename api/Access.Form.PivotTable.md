---
title: Form.PivotTable property (Access)
keywords: vbaac10.chm13521
f1_keywords:
- vbaac10.chm13521
ms.prod: access
api_name:
- Access.Form.PivotTable
ms.assetid: a80edfb5-966b-e1d9-d13e-daefe06c6777
ms.date: 06/08/2017
localization_priority: Normal
---


# Form.PivotTable property (Access)

Returns a  **PivotTable** object representing a PivotTable View on a form. Read-only.


## Syntax

_expression_. `PivotTable`

_expression_ A variable that represents a [Form](Access.Form.md) object.


## Example

This example reports the version of Microsoft Office Web Components in use for the specified form, assuming that there is a PivotTable View on the form.


```vb
Dim objChartSpace As PivotTable 
 
Set objChartSpace = Forms(0).PivotTable 
 
MsgBox "Current version of Office Web Components: " _ 
 & objChartSpace.Version 

```


## See also


[Form Object](Access.Form.md)

