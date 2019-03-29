---
title: Characters object (Excel)
keywords: vbaxl10.chm251072
f1_keywords:
- vbaxl10.chm251072
ms.prod: excel
api_name:
- Excel.Characters
ms.assetid: 128c9ee4-8ba3-6d22-ad0f-9f20be1e24af
ms.date: 03/29/2019
localization_priority: Normal
---


# Characters object (Excel)

Represents characters in an object that contains text. 


## Remarks

The **Characters** object lets you modify any sequence of characters contained in the full text string.

Use **Characters** (_start_, _length_), where _start_ is the start character number and _length_ is the number of characters, to return a **Characters** object.


## Example

The following example adds text to cell B1 and then makes the second word bold.

```vb
With Worksheets("Sheet1").Range("B1") 
 .Value = "New Title" 
 .Characters(5, 5).Font.Bold = True 
End With
```

<br/>

The **[Characters](Excel.Range.Characters.md)** property of the **Range** object is necessary only when you need to change some of an object's text without affecting the rest (you cannot use the **Characters** property to format a portion of the text if the object doesn't support rich text). To change all the text at the same time, you can usually apply the appropriate method or property directly to the object. The following example formats the contents of cell A5 as italic.

```vb
Worksheets("Sheet1").Range("A5").Font.Italic = True
```


## Methods

- [Delete](Excel.Characters.Delete.md)
- [Insert](Excel.Characters.Insert.md)

## Properties

- [Application](Excel.Characters.Application.md)
- [Caption](Excel.Characters.Caption.md)
- [Count](Excel.Characters.Count.md)
- [Creator](Excel.Characters.Creator.md)
- [Font](Excel.Characters.Font.md)
- [Parent](Excel.Characters.Parent.md)
- [PhoneticCharacters](Excel.Characters.PhoneticCharacters.md)
- [Text](Excel.Characters.Text.md)

## See also

- [Excel Object Model Reference](overview/Excel/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
