---
title: Shape.Select method (Publisher)
keywords: vbapb10.chm2228263
f1_keywords:
- vbapb10.chm2228263
ms.prod: publisher
api_name:
- Publisher.Shape.Select
ms.assetid: d18914fd-7679-e922-090c-78affdb39d6a
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.Select method (Publisher)

Selects the specified object.


## Syntax

_expression_.**Select** (_Replace_)

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Replace_|Optional| **Variant**|Specifies whether the selection replaces any previous selection. **True** to replace the previous selection with the new selection; **False** to add the new selection to the previous selection. Default is **True**.|

## Example

This example selects shapes one and three on page one in the active publication.

```vb
ActiveDocument.Pages(1).Shapes.Range(Array(1, 3)).Select
```

<br/>

This example adds shapes two and four on page one in the active publication to the previous selection.

```vb
ActiveDocument.Pages(1).Shapes.Range(Array(2, 4)) _ 
 .Select Replace:=False
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]