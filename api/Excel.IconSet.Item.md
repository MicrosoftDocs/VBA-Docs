---
title: IconSet.Item property (Excel)
keywords: vbaxl10.chm818077
f1_keywords:
- vbaxl10.chm818077
ms.prod: excel
api_name:
- Excel.IconSet.Item
ms.assetid: 4208ddeb-dedb-3d96-c705-adddfcd9a2fe
ms.date: 04/27/2019
localization_priority: Normal
---


# IconSet.Item property (Excel)

Returns an **[Icon](Excel.Icon.md)** object that represents a single icon from an icon set. Read-only.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[IconSet](Excel.IconSet.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index number of the **Icon** object.|

## Remarks

The value of the _Index_ parameter cannot be greater than the number of icons in an icon set. To find the total number of icons in an icon set, use the **[Count](Excel.IconSet.Count.md)** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]