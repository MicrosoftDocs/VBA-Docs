---
title: Phonetic.CharacterType property (Excel)
keywords: vbaxl10.chm628074
f1_keywords:
- vbaxl10.chm628074
ms.prod: excel
api_name:
- Excel.Phonetic.CharacterType
ms.assetid: 2c8ba9b0-1d87-7627-7083-31c9260b68b5
ms.date: 05/03/2019
localization_priority: Normal
---


# Phonetic.CharacterType property (Excel)

Returns or sets the type of phonetic text in the specified cell. Read/write **[XlPhoneticCharacterType](Excel.XlPhoneticCharacterType.md)**.


## Syntax

_expression_.**CharacterType**

_expression_ A variable that represents a **[Phonetic](Excel.Phonetic.md)** object.


## Example

This example changes the first phonetic text string in the active cell from Furigana to Hiragana.

```vb
ActiveCell.Phonetics(1).CharacterType = xlHiragana
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]