---
title: PivotItems.Add method (Excel)
keywords: vbaxl10.chm248074
f1_keywords:
- vbaxl10.chm248074
ms.prod: excel
api_name:
- Excel.PivotItems.Add
ms.assetid: 2d24bb3f-e765-c78c-bef0-787db82056c7
ms.date: 05/07/2019
localization_priority: Normal
---


# PivotItems.Add method (Excel)

Creates a new PivotTable item.


## Syntax

_expression_.**Add** (_Name_)

_expression_ A variable that represents a **[PivotItems](Excel.PivotItems.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new PivotTable item.|

## Example

This example creates a new PivotTable item in the first PivotTable report on worksheet one.

```vb
Worksheets(1).PivotTables(1).PivotItems("Year").Add "1998"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]