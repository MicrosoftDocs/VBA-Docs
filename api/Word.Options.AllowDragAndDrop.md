---
title: Options.AllowDragAndDrop property (Word)
keywords: vbawd10.chm162988100
f1_keywords:
- vbawd10.chm162988100
ms.prod: word
api_name:
- Word.Options.AllowDragAndDrop
ms.assetid: f3cea42e-5fba-7415-bb7a-f214882cc566
ms.date: 06/08/2017
localization_priority: Normal
---


# Options.AllowDragAndDrop property (Word)

 **True** if dragging can be used to move or copy a selection. Read/write **Boolean**.


## Syntax

 _expression_. `AllowDragAndDrop`

 _expression_ A variable that represents an '[Options](Word.Options.md)' object.


## Example

This example turns on the drag-and-drop editing feature.


```vb
Options.AllowDragAndDrop = True
```

This example returns the status of the Drag-and-Drop text-editing option on the Edit tab in the Options dialog box.




```vb
Dim blnDragAndDrop as Boolean 
 
blnDragAndDrop = Options.AllowDragAndDrop
```


## See also


[Options Object](Word.Options.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]