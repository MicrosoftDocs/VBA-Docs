---
title: Options.ArabicMode property (Word)
keywords: vbawd10.chm162988444
f1_keywords:
- vbawd10.chm162988444
ms.prod: word
api_name:
- Word.Options.ArabicMode
ms.assetid: f803708b-2e7d-16bf-5189-07057219c1f0
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.ArabicMode property (Word)

Returns or sets the mode for the Arabic spelling checker. Read/write  **WdAraSpeller**.


## Syntax

 _expression_. `ArabicMode`

 _expression_ Required. A variable that represents an '[Options](Word.Options.md)' collection.


## Example

This example sets the spelling checker to ignore spelling rules regarding Arabic words beginning with an alef hamza.


```vb
Options.ArabicMode = wdInitialAlef
```


## See also


[Options Object](Word.Options.md)

