---
title: ListFormat.ListPictureBullet property (Word)
keywords: vbawd10.chm163577932
f1_keywords:
- vbawd10.chm163577932
ms.prod: word
api_name:
- Word.ListFormat.ListPictureBullet
ms.assetid: b94322ca-ec3a-9aa7-6aa8-db2af124034e
ms.date: 06/08/2017
localization_priority: Normal
---


# ListFormat.ListPictureBullet property (Word)

Returns the **[InlineShape](Word.InlineShape.md)** object that represents the picture used as a bullet in a picture bulleted list.


## Syntax

_expression_. `ListPictureBullet`

 _expression_ An expression that returns a '[ListFormat](Word.ListFormat.md)' object.


## Example

This example sets the height and width of the selected picture bullet. This example assumes that the insertion point in the document is located in a paragraph formatted with a picture bullet.


```vb
Sub ListPictBullet() 
 With Selection.Range.ListFormat.ListPictureBullet 
 .Width = InchesToPoints(Inches:=0.5) 
 .Height = InchesToPoints(Inches:=0.05) 
 End With 
End Sub
```


## See also


[ListFormat Object](Word.ListFormat.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]