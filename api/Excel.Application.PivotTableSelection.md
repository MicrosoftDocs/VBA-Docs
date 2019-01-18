---
title: Application.PivotTableSelection property (Excel)
keywords: vbaxl10.chm133192
f1_keywords:
- vbaxl10.chm133192
ms.prod: excel
api_name:
- Excel.Application.PivotTableSelection
ms.assetid: e0a93c11-2e2f-23af-6cad-b4f22883128e
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.PivotTableSelection property (Excel)

 **True** if PivotTable reports use structured selection. Read/write **Boolean**.


## Syntax

_expression_. `PivotTableSelection`

_expression_ A variable that represents an [Application](Excel.Application-graph-property.md) object.


## Example

This example enables structured selection mode and then sets the first PivotTable report on worksheet one to allow only data to be selected.


```vb
Application.PivotTableSelection = True 
Worksheets(1).PivotTables(1).SelectionMode = xlDataOnly
```


## See also


[Application Object](Excel.Application(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]