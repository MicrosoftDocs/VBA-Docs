---
title: Selection.GoToPrevious method (Word)
keywords: vbawd10.chm158662831
f1_keywords:
- vbawd10.chm158662831
ms.prod: word
api_name:
- Word.Selection.GoToPrevious
ms.assetid: da41b0b4-673e-5701-d31d-ab3314600e53
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.GoToPrevious method (Word)

Returns a  **Range** object that refers to the start position of the previous item or location specified by the What argument. If applied to a **Selection** object, **GoToPrevious** moves the selection to the specified item. **Range** object.


## Syntax

_expression_. `GoToPrevious`( `_What_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _What_|Required| **WdGoToItem**|The item where the specified range or selection is to be moved.|

## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]