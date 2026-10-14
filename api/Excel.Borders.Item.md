---
title: Borders.Item property (Excel)
keywords: vbaxl10.chm181076
f1_keywords:
- vbaxl10.chm181076
api_name:
- Excel.Borders.Item
ms.assetid: 19184379-d551-396e-8cb6-ff240e3c85fa
ms.date: 10/12/2026
ms.localizationpriority: medium
---


# Borders.Item property (Excel)

Returns a **[Border](Excel.Border(object).md)** object that represents one of the borders of a range of cells, a style, or a conditional format.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Borders](Excel.Borders.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **[XlBordersIndex](excel.xlbordersindex.md)**|The border to return. The supported values depend on the object that returned the **Borders** collection. See the Remarks section.|


## Remarks

The supported value of _Index_ depends on the object that returned the **Borders** collection:

- For **[Range.Borders](Excel.Range.Borders.md)** and **[Style.Borders](Excel.Style.Borders.md)**, use the applicable **XlBordersIndex** constants, such as **xlEdgeBottom**.
- For **[FormatCondition.Borders](Excel.FormatCondition.Borders.md)**, use **xlLeft**, **xlRight**, **xlTop**, or **xlBottom**. Don't use the corresponding **xlEdge** constants.

## Example

This following example sets the color of the bottom border of cells A1:G1.

```vb
Worksheets("Sheet1").Range("a1:g1"). _ 
 Borders.Item(xlEdgeBottom).Color = RGB(255, 0, 0)
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]