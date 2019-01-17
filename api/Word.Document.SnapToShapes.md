---
title: Document.SnapToShapes property (Word)
keywords: vbawd10.chm158007597
f1_keywords:
- vbawd10.chm158007597
ms.prod: word
api_name:
- Word.Document.SnapToShapes
ms.assetid: b74e7a58-deee-aed2-8956-3911dd54d9ba
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.SnapToShapes property (Word)

 **True** if Microsoft Word automatically aligns AutoShapes or East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other AutoShapes or East Asian characters in the specified document. Read/write **Boolean**.


## Syntax

 _expression_. `SnapToShapes`

 _expression_ A variable that represents a '[Document](Word.Document.md)' object.


## Remarks

This property creates additional invisible gridlines for each AutoShape.  **SnapToShapes** works independently of the **[SnapToGrid](Word.Document.SnapToGrid.md)** property.


## Example

This example sets Microsoft Word to automatically align East Asian characters with invisible gridlines that go through the vertical and horizontal edges of other East Asian characters in the current document.


```vb
ActiveDocument.SnapToShapes = True
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]