---
title: Frame object (Word)
keywords: vbawd10.chm2346
f1_keywords:
- vbawd10.chm2346
ms.prod: word
api_name:
- Word.Frame
ms.assetid: d36d3361-9e93-7dd9-b8c9-0ce503e03810
ms.date: 06/08/2017
localization_priority: Normal
---


# Frame object (Word)

Represents a frame. The **Frame** object is a member of the **Frames** collection. The **[Frames](Word.Frames.md)** collection includes all frames in a selection, range, or document.


## Remarks

Use  **Frames** (Index), where Index is the index number, to return a single **Frame** object. The index number represents the position of the frame in the selection, range, or document. The following example allows text to wrap around the first frame in the active document.


```vb
ActiveDocument.Frames(1).TextWrap = True
```

Use the **Add** method to add a frame around a range. The following example adds a frame around the first paragraph in the active document.




```vb
ActiveDocument.Frames.Add _ 
 Range:=ActiveDocument.Paragraphs(1).Range
```

You can wrap text around  **Shape** or **ShapeRange** objects by using the **WrapFormat** property. You can position a **Shape** or **ShapeRange** object by using the **Top** and **Left** properties.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]