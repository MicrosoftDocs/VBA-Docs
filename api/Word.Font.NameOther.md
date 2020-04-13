---
title: Font.NameOther property (Word)
keywords: vbawd10.chm156369054
f1_keywords:
- vbawd10.chm156369054
ms.prod: word
api_name:
- Word.Font.NameOther
ms.assetid: d3bfd1f6-e561-ed05-b0a6-e886d6e2264c
ms.date: 06/08/2017
localization_priority: Normal
---


# Font.NameOther property (Word)

Returns or sets the font used for characters with character codes from 128 through 255. Read/write  **String**.


## Syntax

_expression_. `NameOther`

 _expression_ An expression that returns a **[Font](Word.Font.md)** object.


## Remarks

In the U.S. English version of Microsoft Word, the default value of this property is Times New Roman. Use the **[Name](Word.Font.Name.md)** property to change the font that's applied to all text and that appears in the **Font** box on the **Formatting** toolbar.


## Example

This example sets the font used for characters with character codes from 128 through 255.


```vb
Selection.Font.NameOther = "Century"
```


## See also


[Font Object](Word.Font.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]