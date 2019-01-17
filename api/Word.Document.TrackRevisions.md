---
title: Document.TrackRevisions property (Word)
keywords: vbawd10.chm158007610
f1_keywords:
- vbawd10.chm158007610
ms.prod: word
api_name:
- Word.Document.TrackRevisions
ms.assetid: c6ff8462-805d-2494-cebb-ace6fe536f40
ms.date: 06/08/2017
localization_priority: Normal
---


# Document.TrackRevisions property (Word)

 **True** if changes are tracked in the specified document. Read/write **Boolean**.


## Syntax

 _expression_. `TrackRevisions`

 _expression_ A variable that represents a '[Document](Word.Document.md)' object.


## Example

This example sets the active document so that it tracks changes and makes them visible on the screen.


```vb
With ActiveDocument 
 .TrackRevisions = True 
 .ShowRevisions = True 
End With
```

This example inserts text if change tracking isn't enabled.




```vb
If ActiveDocument.TrackRevisions = False Then 
 Selection.InsertBefore "new text" 
End If
```


## See also


[Document Object](Word.Document.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]