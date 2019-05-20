---
title: XPath.Clear method (Excel)
keywords: vbaxl10.chm760077
f1_keywords:
- vbaxl10.chm760077
ms.prod: excel
api_name:
- Excel.XPath.Clear
ms.assetid: 8d9e0c70-c77e-257f-6ac7-7a8577282ab1
ms.date: 05/21/2019
localization_priority: Normal
---


# XPath.Clear method (Excel)

Clears all XPath schema information for the mapped range. 


## Syntax

_expression_.**Clear**

_expression_ A variable that represents an **[XPath](Excel.XPath.md)** object.


## Remarks

**Clear** affects the entire range mapped to the **XPath** object.

This method does not clear the data from the cells mapped to the specified XPath. Use the **[Clear](Excel.Range.Clear.md)** method of the **Range** object to clear the data from the cells.

If the specified XPath is mapped in an XML list, the schema mapping is removed, but the list is not deleted from the worksheet.

If the mapped range is a single-cell, the single-cell is removed and the data remains.

This method produces an error if any of the following conditions are true:

- The range spans multiple columns in the grid.
    
- Part of the range spans already mapped cells and the rest spans unmapped cells.
    
- Part of the range spans one mapping, and another part of the range spans a different mapping or different XPath from the same mapping.
    



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]