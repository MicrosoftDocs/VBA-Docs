---
title: TextStyles.Item method (Publisher)
keywords: vbapb10.chm5898240
f1_keywords:
- vbapb10.chm5898240
ms.prod: publisher
api_name:
- Publisher.TextStyles.Item
ms.assetid: 14d1871f-c2cb-31af-e22d-10b3cf59b6fc
ms.date: 06/15/2019
localization_priority: Normal
---


# TextStyles.Item method (Publisher)

Returns an individual object in a specified collection.


## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[TextStyles](Publisher.TextStyles.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Index_|Required| **Variant**|The number or name of the field or list box item to return.|

## Return value

**[TextStyle](Publisher.TextStyle.md)**


## Example

This example returns the Normal text style from the active publication.

```vb
Dim txtStyle As TextStyle 
 
Set txtStyle = ActiveDocument.TextStyles.Item(Index:="Normal")
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]