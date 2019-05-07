---
title: Document.AutoFormatOverride property (Word)
keywords: vbawd10.chm158007768
f1_keywords:
- vbawd10.chm158007768
ms.prod: word
api_name:
- Word.Document.AutoFormatOverride
ms.assetid: 85287164-98f8-fd3a-36b7-b03008e9aac3
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.AutoFormatOverride property (Word)

Returns or sets a  **Boolean** that represents whether automatic formatting options override formatting restrictions in a document where formatting restrictions are in effect.


## Syntax

_expression_. `AutoFormatOverride`

_expression_ A variable that represents a **[Document](Word.Document.md)** object.


## Example

The following specifies that automatic formatting will override formatting restrictions in a protected document.


```vb
ActiveDocument.AutoFormatOverride = True
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]