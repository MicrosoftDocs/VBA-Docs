---
title: Phonetics.CharacterType property (Excel)
keywords: vbaxl10.chm658077
f1_keywords:
- vbaxl10.chm658077
ms.prod: excel
api_name:
- Excel.Phonetics.CharacterType
ms.assetid: b61c3bd5-86dc-baed-e47f-62d522fca290
ms.date: 05/03/2019
localization_priority: Normal
---


# Phonetics.CharacterType property (Excel)

Returns or sets the type of phonetic text in the specified cell. Read/write **[XlPhoneticCharacterType](Excel.XlPhoneticCharacterType.md)**.


## Syntax

_expression_.**CharacterType**

_expression_ A variable that represents a **[Phonetics](Excel.Phonetics.md)** object.


## Example

This example changes the first phonetic text string in the active cell from Furigana to Hiragana.

```vb
ActiveCell.Phonetics(1).CharacterType = xlHiragana
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]