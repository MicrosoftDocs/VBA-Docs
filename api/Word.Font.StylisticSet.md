---
title: Font.StylisticSet property (Word)
keywords: vbawd10.chm156369074
f1_keywords:
- vbawd10.chm156369074
ms.prod: word
api_name:
- Word.Font.StylisticSet
ms.assetid: e82013b1-9f55-d17a-a510-6f77b627382b
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.StylisticSet property (Word)

Specifies the stylistic set for the specified font. Read/write [WdStylisticSet](Word.WdStylisticSet.md).


## Syntax

_expression_. `StylisticSet`

 _expression_ An expression that returns a **[Font](Word.Font.md)** object.


## Remarks

Some OpenType fonts provide stylistic sets. A stylistic set defines a set of characters within the font that are intended to be used together, usually for the purpose of visual harmony, such as in headings.


## Example

The following code example sets the font for the active document to Gabriola and then applies the sixth stylistic set provided by the Gabriola font.


```vb
ActiveDocument.Range.Font.Name = "Gabriola" 
ActiveDocument.Range.Font.StylisticSet = wdStylisticSet06
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]