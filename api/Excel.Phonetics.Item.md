---
title: Phonetics.Item property (Excel)
keywords: vbaxl10.chm658080
f1_keywords:
- vbaxl10.chm658080
ms.prod: excel
api_name:
- Excel.Phonetics.Item
ms.assetid: 41c2df73-fb88-fe1a-a4ff-4562441b1510
ms.date: 05/03/2019
localization_priority: Normal
---


# Phonetics.Item property (Excel)

Returns a single object from a collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Phonetics](Excel.Phonetics.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Long**|The index number of the object.|

## Example

This example makes the first phonetic text string in the active cell visible.

```vb
ActiveCell.Phonetics.Item(1).Visible = True
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]