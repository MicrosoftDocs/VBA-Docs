---
title: Document.FarEastLineBreakLanguage property (Word)
keywords: vbawd10.chm158007622
f1_keywords:
- vbawd10.chm158007622
ms.prod: word
api_name:
- Word.Document.FarEastLineBreakLanguage
ms.assetid: cf868676-b880-46e9-a1b4-9cb341c63427
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.FarEastLineBreakLanguage property (Word)

Returns or sets a  **[WdFarEastLineBreakLanguageID](Word.WdFarEastLineBreakLanguageID.md)** that represents the East Asian language to use when breaking lines of text in the specified document or template. Read/write.


## Syntax

 _expression_. `reFarEastLineBakLanguage`

 _expression_ Required. A variable that represents a '[Document](Word.Document.md)' object.


## Example

This example sets Word to break lines in the current document based on Korean language rules.


```vb
ActiveDocument.FarEastLineBreakLanguage = wdLineBreakKorean
```


## See also


[Document Object](Word.Document.md)

