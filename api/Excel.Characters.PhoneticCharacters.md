---
title: Characters.PhoneticCharacters Property (Excel)
keywords: vbaxl10.chm252079
f1_keywords:
- vbaxl10.chm252079
ms.prod: excel
api_name:
- Excel.Characters.PhoneticCharacters
ms.assetid: 05e5cfa5-aef8-c413-29e4-3c608bd4f953
ms.date: 06/08/2017
---


# Characters.PhoneticCharacters Property (Excel)

Returns or sets the phonetic text in the specified  **[Characters](Excel.Characters.md)** object. Read/write **String** .


## Syntax

 _expression_ . **PhoneticCharacters**

 _expression_ A variable that represents a **Characters** object.


## Remarks

Instead of using this property, you should use the  **[Add](Excel.Phonetics.Add.md)** method of the **[Phonetics](Excel.Phonetics.md)** collection to add phonetic information to a cell, and use the **[Text](Excel.Phonetic.Text.md)** property of the **[Phonetic](Excel.Phonetic.md)** object to return or set the phonetic text strings in a cell.


## Example

This example replaces the fourth character from the beginning of the text in the active cell with Furigana characters.


```vb
ActiveCell.Characters(1,3).PhoneticCharacters = "フリガナ"
```


## See also


#### Concepts


[Characters Object](Excel.Characters.md)

