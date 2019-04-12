---
title: Borders.Item property (Excel)
keywords: vbaxl10.chm181076
f1_keywords:
- vbaxl10.chm181076
ms.prod: excel
api_name:
- Excel.Borders.Item
ms.assetid: 19184379-d551-396e-8cb6-ff240e3c85fa
ms.date: 04/13/2019
localization_priority: Normal
---


# Borders.Item property (Excel)

Returns a **[Border](Excel.Border(object).md)** object that represents one of the borders of either a range of cells or a style.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Borders](Excel.Borders.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **[XlBordersIndex](excel.xlbordersindex.md)**|Can be one of the **XlBordersIndex** constants.|


## Example

This following example sets the color of the bottom border of cells A1:G1.

```vb
Worksheets("Sheet1").Range("a1:g1"). _ 
 Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]