---
title: Style.Font property (Word)
keywords: vbawd10.chm153878538
f1_keywords:
- vbawd10.chm153878538
ms.prod: word
api_name:
- Word.Style.Font
ms.assetid: e4e5968a-ab2e-786b-cc71-f770d8c121b4
ms.date: 06/08/2017
localization_priority: Normal
---


# Style.Font property (Word)

Returns or sets a  **[Font](Word.Font.md)** object that represents the character formatting of the specified style. Read/write **Font**.


## Syntax

 _expression_. `Font`

 _expression_ A variable that represents a '[Style](Word.Style.md)' object.


## Remarks

To set this property, specify an expression that returns a  **[Font](Word.Font.md)** object.


## Example

This example removes bold formatting from the Heading 1 style in the active document.


```vb
ActiveDocument.Styles(wdStyleHeading1).Font.Bold = False
```


## See also


[Style Object](Word.Style.md)

