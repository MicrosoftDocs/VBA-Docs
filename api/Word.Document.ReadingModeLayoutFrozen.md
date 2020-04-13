---
title: Document.ReadingModeLayoutFrozen property (Word)
keywords: vbawd10.chm158007777
f1_keywords:
- vbawd10.chm158007777
ms.prod: word
api_name:
- Word.Document.ReadingModeLayoutFrozen
ms.assetid: 5ca8aef3-82dd-81c6-9620-57f304bcbb64
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReadingModeLayoutFrozen property (Word)

Sets or returns a  **Boolean** that represents whether pages displayed in reading layout view are frozen to a specified size for inserting handwritten markup into a document.


## Syntax

_expression_. `ReadingModeLayoutFrozen`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Remarks

Use the **[ReadingLayoutSizeX](Word.Document.ReadingLayoutSizeX.md)** and **[ReadingLayoutSizeY](Word.Document.ReadingLayoutSizeY.md)** properties to specify the size of the pages displayed when the reading layout size is frozen for inserting handwritten markup into a document.


## Example

The following example displays the active document in reading layout view and then sets the size of the displayed pages.


```vb
ActiveWindow.View.ReadingLayout = True 
ActiveDocument.ReadingLayoutSize 300, 300 
ActiveDocument.ReadingModeLayoutFrozen = True
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]