---
title: Selection.GoToNext method (Word)
keywords: vbawd10.chm158662830
f1_keywords:
- vbawd10.chm158662830
ms.prod: word
api_name:
- Word.Selection.GoToNext
ms.assetid: af6a4e91-7ec1-929a-7577-4e457f5ce1bd
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.GoToNext method (Word)

Returns a  **Range** object that refers to the start position of the next item or location specified by the What argument. If you apply this method to the **Selection** object, the method moves the selection to the specified item (except for the **wdGoToGrammaticalError**, **wdGoToProofreadingError**, and **wdGoToSpellingError** constants).


## Syntax

_expression_. `GoToNext`( `_What_` )

_expression_ Required. A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _What_|Required| **WdGoToItem**|The item where the specified range or selection is to be moved.|

## Remarks




> [!NOTE] 
> When you use this method with the **wdGoToGrammaticalError**, **wdGoToProofreadingError**, or **wdGoToSpellingError** constant, the **Range** object that's returned includes any grammar error text or spelling error text.


## See also


[Selection Object](Word.Selection.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]