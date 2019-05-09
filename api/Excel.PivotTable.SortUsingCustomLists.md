---
title: PivotTable.SortUsingCustomLists property (Excel)
keywords: vbaxl10.chm235182
f1_keywords:
- vbaxl10.chm235182
ms.prod: excel
api_name:
- Excel.PivotTable.SortUsingCustomLists
ms.assetid: ff7a8a4d-9d64-f6dd-c373-e979d016f741
ms.date: 05/09/2019
localization_priority: Normal
---


# PivotTable.SortUsingCustomLists property (Excel)

The **SortUsingCustomLists** property controls whether custom lists are used for sorting items of fields, both initially when the PivotField is initialized and the PivotItems are ordered by their captions, and later when the user applies a sort. Read/write **Boolean**.


## Syntax

_expression_.**SortUsingCustomLists**

_expression_ A variable that represents a **[PivotTable](Excel.PivotTable.md)** object.


## Remarks

Setting this property to **False** can optimize performance for fields with many items, and it also allows users that do not want custom list-based sorting to avoid it.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]