---
title: Range.CountLarge property (Excel)
keywords: vbaxl10.chm144247
f1_keywords:
- vbaxl10.chm144247
ms.prod: excel
api_name:
- Excel.Range.CountLarge
ms.assetid: 3a46ef6d-a339-b15e-990d-b11f462fb602
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.CountLarge property (Excel)

Returns a value that represents the number of objects in the collection. Read-only **Variant**.


## Syntax

_expression_.**CountLarge**

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Remarks

The **CountLarge** property is functionally the same as the **[Count](Excel.Range.Count.md)** property, except that the **Count** property will generate an overflow error if the specified range has more than 2,147,483,647 cells (one less than 2,048 columns). The **CountLarge** property, however, can handle ranges up to the maximum size for a worksheet, which is 17,179,869,184 cells.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
