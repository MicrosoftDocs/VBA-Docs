---
title: Range.XPath property (Excel)
keywords: vbaxl10.chm144241
f1_keywords:
- vbaxl10.chm144241
ms.prod: excel
api_name:
- Excel.Range.XPath
ms.assetid: 90a353d7-7222-b387-558a-044cb17f09b9
ms.date: 05/11/2019
localization_priority: Normal
---


# Range.XPath property (Excel)

Returns an **[XPath](Excel.XPath.md)** object that represents the XPath of the element mapped to the specified **Range** object. The context of the range determines whether the action succeeds or returns an empty object. Read-only.


## Syntax

_expression_.**XPath**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The **XPath** property is valid when the range it contains meets the following conditions:

- The range is a single cell.
    
- If the range consists of two or more cells, one or the other must be true:
    
  - If the cells contain XPath information, all cells in the range must contain XPath information (that is, each cell is associated with one or more data maps), and all of the cells must have identical XPath content (that is, each cell contributes to the same set of data maps).
    
  - All of the cells must contain no XPath information.
    
- The range does not contain discontinuous areas.
    
> [!NOTE] 
> The header and totals row of a table are considered to contain XPath information. Any ranges that don't meet the above conditions returns a run-time error.

If the range selection is valid, but none of the cells are mapped, Excel returns an **XPath** object so that you can access the **SetValue** method to create a mapping.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]