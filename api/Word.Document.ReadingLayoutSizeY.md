---
title: Document.ReadingLayoutSizeY property (Word)
keywords: vbawd10.chm158007788
f1_keywords:
- vbawd10.chm158007788
ms.prod: word
api_name:
- Word.Document.ReadingLayoutSizeY
ms.assetid: dc2f437c-56cd-9bd6-5808-4489e48e5b90
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.ReadingLayoutSizeY property (Word)

Sets or returns a  **Long** that represents the height of pages in a document when it is displayed in reading layout view and is frozen for entering handwritten markup.


## Syntax

_expression_. `ReadingLayoutSizeY`

 _expression_ An expression that returns a **[Document](Word.Document.md)** object.


## Remarks

After setting the **[ReadingLayoutSizeX](Word.Document.ReadingLayoutSizeX.md)** and **ReadingLayoutSizeY** properties, use the **[ReadingModeLayoutFrozen](Word.Document.ReadingModeLayoutFrozen.md)** property to display the page using the specified height and width. Use the **[ReadingLayout](Word.View.ReadingLayout.md)** property to display a document in reading layout view.


## Example

The following example displays the active document in reading layout view, and then sets the size of the displayed pages.


```vb
ActiveWindow.View.ReadingLayout = True 
ActiveDocument.ReadingLayoutSizeX = 300 
ActiveDocument.ReadingLayoutSizeY = 300 
ActiveDocument.ReadingModeLayoutFrozen = True
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]