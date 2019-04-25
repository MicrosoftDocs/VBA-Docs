---
title: HPageBreaks.Item property (Excel)
keywords: vbaxl10.chm165073
f1_keywords:
- vbaxl10.chm165073
ms.prod: excel
api_name:
- Excel.HPageBreaks.Item
ms.assetid: 2c216336-ed46-382b-e408-3de708afb3c3
ms.date: 04/26/2019
localization_priority: Normal
---


# HPageBreaks.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents an **[HPageBreaks](Excel.HPageBreaks.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example changes the location of horizontal page break one.

```vb
Worksheets(1).HPageBreaks.Item(1).Location = .Range("e5")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]