---
title: Characters.PhoneticCharacters property (Excel)
keywords: vbaxl10.chm252079
f1_keywords:
- vbaxl10.chm252079
ms.prod: excel
api_name:
- Excel.Characters.PhoneticCharacters
ms.assetid: 05e5cfa5-aef8-c413-29e4-3c608bd4f953
ms.date: 04/16/2019
localization_priority: Normal
---


# Characters.PhoneticCharacters property (Excel)

Returns or sets the phonetic text in the specified **Characters** object. Read/write **String**.


## Syntax

_expression_.**PhoneticCharacters**

_expression_ A variable that represents a **[Characters](Excel.Characters.md)** object.


## Remarks

Instead of using this property, you should use the **[Add](Excel.Phonetics.Add.md)** method of the **Phonetics** collection to add phonetic information to a cell, and use the **[Text](Excel.Phonetic.Text.md)** property of the **Phonetic** object to return or set the phonetic text strings in a cell.


## Example

This example replaces the fourth character from the beginning of the text in the active cell with Furigana characters.

```vb
ActiveCell.Characters(1,3).PhoneticCharacters = "フリガナ"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]