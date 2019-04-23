---
title: ChartGroups.Item method (Excel)
keywords: vbaxl10.chm570074
f1_keywords:
- vbaxl10.chm570074
ms.prod: excel
api_name:
- Excel.ChartGroups.Item
ms.assetid: 29ca6f13-96b7-bd43-9562-480c467ef7db
ms.date: 04/20/2019
localization_priority: Normal
---


# ChartGroups.Item method (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[ChartGroups](Excel.ChartGroups(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number for the object.|

## Return value

A **[ChartGroup](Excel.ChartGroup(object).md)** object contained by the collection.


## Example

This example adds drop lines to chart group one on chart sheet one.

```vb
Charts(1).ChartGroups.Item(1).HasDropLines = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]