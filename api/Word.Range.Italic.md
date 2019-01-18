---
title: Range.Italic property (Word)
keywords: vbawd10.chm157155459
f1_keywords:
- vbawd10.chm157155459
ms.prod: word
api_name:
- Word.Range.Italic
ms.assetid: 7d52781a-46f2-7bca-067e-dc41772149fc
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.Italic property (Word)

 **True** if the font or range is formatted as italic. Read/write **Long**.


## Syntax

 _expression_. `Italic`

 _expression_ A variable that represents a '[Range](Word.Range.md)' object.


## Remarks

This property returns  **True** , **False** or **wdUndefined** (a mixture of **True** and **False**) and can be set to **True** , **False** , or **wdToggle**.


## Example

This example formats the first word in the active document as italic.


```vb
ActiveDocument.Words(1).Italic = True
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]