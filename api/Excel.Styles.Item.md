---
title: Styles.Item property (Excel)
keywords: vbaxl10.chm179075
f1_keywords:
- vbaxl10.chm179075
ms.prod: excel
api_name:
- Excel.Styles.Item
ms.assetid: 2101cf1a-b37f-23f8-25b2-dde124d7c702
ms.date: 05/16/2019
localization_priority: Normal
---


# Styles.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Styles](Excel.Styles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or index number of the object.|

## Example

This example changes the Normal style for the active workbook by setting the style's **Bold** property.

```vb
ActiveWorkbook.Styles.Item("Normal").Font.Bold = True
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]