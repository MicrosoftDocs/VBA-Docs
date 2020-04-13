---
title: Range.GoToNext method (Word)
keywords: vbawd10.chm157155502
f1_keywords:
- vbawd10.chm157155502
ms.prod: word
api_name:
- Word.Range.GoToNext
ms.assetid: 011de2d6-c0fc-608f-8d7e-faac5947978d
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.GoToNext method (Word)

Returns a  **Range** object that refers to the start position of the next item or location specified by the What argument. .


## Syntax

_expression_. `GoToNext`( `_What_` )

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _What_|Required| **[WdGoToItem](Word.WdGoToItem.md)**|The item to where the specified range or selection is to be moved.|

## Remarks




> [!NOTE] 
> When you use this method with the **wdGoToGrammaticalError**, **wdGoToProofreadingError**, or **wdGoToSpellingError** constant, the **Range** object that's returned includes any grammar error text or spelling error text.


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]