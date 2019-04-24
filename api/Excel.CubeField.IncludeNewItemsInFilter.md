---
title: CubeField.IncludeNewItemsInFilter property (Excel)
keywords: vbaxl10.chm668095
f1_keywords:
- vbaxl10.chm668095
ms.prod: excel
api_name:
- Excel.CubeField.IncludeNewItemsInFilter
ms.assetid: 7c9ccb66-5a8c-ced0-c024-2336e85f00db
ms.date: 04/23/2019
localization_priority: Normal
---


# CubeField.IncludeNewItemsInFilter property (Excel)

The **IncludeNewItemsInFilter** property is used to track included/excluded items in OLAP PivotTables. Read/write.


## Syntax

_expression_.**IncludeNewItemsInFilter**

_expression_ A variable that represents a **[CubeField](Excel.CubeField.md)** object.


## Remarks

The default value is **False**.

When this setting is set to **True**, excluded items are tracked when manual filtering is applied. When this setting is set to **False**, included items are tracked when manual filtering is applied.

When **IncludeNewItemsInFilter** is set to **False**, the **HiddenItemsList** and **HiddenItems** collections are empty, and items cannot be added to them.

When **IncludeNewItemsInFilter** is set to **True**, the **VisibleItemsList** and **VisibleItems** collections are empty, and items cannot be added to them.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]