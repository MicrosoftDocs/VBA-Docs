---
title: Revisions object (Word)
ms.prod: word
ms.assetid: 7f267a64-885a-cb4c-008a-e8545cea94d2
ms.date: 06/08/2017
localization_priority: Normal
---


# Revisions object (Word)

A collection of  **[Revision](Word.Revision.md)** objects that represent the changes marked with revision marks in a range or document.


## Remarks

Use the  **Revisions** property to return the **Revisions** collection. The following code example displays the number of revisions in the main text story.


```vb
MsgBox ActiveDocument.Revisions.Count
```

The following code example accepts all the revisions in the selection.




```vb
For Each myRev In Selection.Range.Revisions 
 myRev.Accept 
Next myRev
```

The following code example accepts all the revisions in the first paragraph in the selection.




```vb
Set myRange = Selection.Paragraphs(1).Range 
myRange.Revisions.AcceptAll
```

The  **Add** method is not available for the **Revisions** collection. **Revision** objects are added when change tracking is enabled. Set the **TrackRevisions** property to **True** to track revisions made to the document text. The following code example enables revision tracking in the active document and then inserts "The " before the selection.




```vb
ActiveDocument.TrackRevisions = True 
Selection.InsertBefore "The "
```

Use  **Revisions** (Index), where Index is the index number, to return a single **Revision** object. The index number represents the position of the revision in the range or document. The following code example displays the author name for the first revision in the first section.




```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions(1).Author
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## Methods



|Name|
|:-----|
|[AcceptAll](Word.Revisions.AcceptAll.md)|
|[Item](Word.Revisions.Item.md)|
|[RejectAll](Word.Revisions.RejectAll.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Revisions.Application.md)|
|[Count](Word.Revisions.Count.md)|
|[Creator](Word.Revisions.Creator.md)|
|[Parent](Word.Revisions.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
