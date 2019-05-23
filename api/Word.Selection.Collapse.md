---
title: Selection.Collapse method (Word)
keywords: vbawd10.chm158662757
f1_keywords:
- vbawd10.chm158662757
ms.prod: word
api_name:
- Word.Selection.Collapse
ms.assetid: 92ccd3dc-41ab-b3d4-5397-fca7d7f01635
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.Collapse method (Word)

Collapses a selection to the starting or ending position. After a selection is collapsed, the starting and ending points are equal.


## Syntax

_expression_. `Collapse`( `_Direction_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Direction_|Optional| **Variant**|The direction in which to collapse the range or selection. Can be either of the following  **WdCollapseDirection** constants: **wdCollapseEnd** or **wdCollapseStart**. The default value is **wdCollapseStart**.|

## Example

This example collapses the selection to an insertion point at the beginning of the previous selection.


```vb
Selection.Collapse Direction:=wdCollapseStart
```


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
