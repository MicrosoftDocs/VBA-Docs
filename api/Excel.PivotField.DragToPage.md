---
title: PivotField.DragToPage property (Excel)
keywords: vbaxl10.chm240104
f1_keywords:
- vbaxl10.chm240104
ms.prod: excel
api_name:
- Excel.PivotField.DragToPage
ms.assetid: 3bca0805-8f9f-099a-cd9f-3621025654e5
ms.date: 05/04/2019
localization_priority: Normal
---


# PivotField.DragToPage property (Excel)

**True** if the field can be dragged to the page position. The default value is **True**. Read/write **Boolean**.


## Syntax

_expression_.**DragToPage**

_expression_ A variable that represents a **[PivotField](Excel.PivotField.md)** object.


## Remarks

For OLAP data sources, the value is **False** for measure fields.


## Example

This example prevents the Year field in the PivotTable report on worksheet one from being dragged to the page position.

```vb
Worksheets(1).PivotTables("Pivot1") _ 
 .PivotFields("Year").DragToPage = False
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]