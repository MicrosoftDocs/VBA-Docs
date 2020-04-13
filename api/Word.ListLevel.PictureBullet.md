---
title: ListLevel.PictureBullet property (Word)
keywords: vbawd10.chm160235534
f1_keywords:
- vbawd10.chm160235534
ms.prod: word
api_name:
- Word.ListLevel.PictureBullet
ms.assetid: 73c44f47-182c-9ef6-106c-fd68000a27c3
ms.date: 06/08/2017
localization_priority: Normal
---


# ListLevel.PictureBullet property (Word)

Returns an  **[InlineShape](Word.InlineShape.md)** object that represents a picture bullet.


## Syntax

_expression_. `PictureBullet`

 _expression_ An expression that returns a '[ListLevel](Word.ListLevel.md)' object.


## Example

This example returns the picture bullet for the first list in the active document and sets the picture bullet's width to 0.25 inch. To see this example, first run the code example for the **[ApplyPictureBullet](Word.ListLevel.ApplyPictureBullet.md)** method.


```vb
Sub PicBullet() 
 ActiveDocument.ListTemplates(1) _ 
 .ListLevels(1) _ 
 .PictureBullet.Width = InchesToPoints(0.25) 
End Sub
```


## See also


[ListLevel Object](Word.ListLevel.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]