---
title: CoAuthUpdate object (Word)
ms.prod: word
api_name:
- Word.CoAuthUpdate
ms.assetid: c00e5029-2e4b-97c0-33d3-86fdc53df535
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthUpdate object (Word)

Represents a range of text that has been updated by a co author.


## Remarks

When a document that has co authoring enabled is edited by more than one author, changes to the document by one author are pushed to other authors' versions of the document by using updates. When a co author performs an explicit document save (by pressing  **CTRL** + **S**, for example), changes made by other co authors are merged into the document as updates. The **CoAuthUpdates** collection contains all changes that were merged into the document, where each change is a single update represented by a **CoAuthUpdate** object.

The contents of the **CoAuthUpdates** collection remains the same until a co author performs another explicit document save. When the co author saves the document again, if there are no new changes from other co authors that are merged into the document, the **CoAuthUpdates** collection retains the same updates that were merged at the previous explicit save. If there are new changes that are merged into the document, the **CoAuthUpdates** collection contains the new updates for the document. Use a **CoAuthUpdate** object to retrieve an individual update from the **[CoAuthUpdates](overview/Word.md)** collection.


## Example

The following code example gets the associated text in the range of each  **CoAuthUpdate** object in the active document.


```vb
Dim caUpdate As CoAuthUpdate 
Dim strText As String 
 
For Each caUpdate In ActiveDocument.CoAuthoring.Updates 
    strText = caUpdate.Range.Text 
    MsgBox strText 
Next caUpdate
```


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]