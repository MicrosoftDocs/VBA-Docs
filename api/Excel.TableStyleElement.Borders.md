---
title: TableStyleElement.Borders property (Excel)
keywords: vbaxl10.chm835075
f1_keywords:
- vbaxl10.chm835075
ms.prod: excel
api_name:
- Excel.TableStyleElement.Borders
ms.assetid: a6fdfe85-0953-f796-5c89-6f418e9226e6
ms.date: 05/17/2019
localization_priority: Normal
---


# TableStyleElement.Borders property (Excel)

Returns a **[Borders](Excel.Borders.md)** collection that represents the borders of a **TableStyleElement** object. Read-only.


## Syntax

_expression_.**Borders**

_expression_ A variable that represents a **[TableStyleElement](Excel.TableStyleElement.md)** object.


## Example

This example sets the color of the top border of a table to red.

```vb
With ActiveWorkbook.TableStyles("Table Style 4").TableStyleElements( _ 
 xlWholeTable).Borders(xlEdgeTop) 
 .Color = 255 
 .TintAndShade = 0 
 .Weight = 2 
 .LineStyle = 1 
End With
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]