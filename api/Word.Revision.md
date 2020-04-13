---
title: Revision object (Word)
keywords: vbawd10.chm2433
f1_keywords:
- vbawd10.chm2433
ms.prod: word
api_name:
- Word.Revision
ms.assetid: e6f64467-a438-88f1-60f9-975365a1430e
ms.date: 06/08/2017
localization_priority: Normal
---


# Revision object (Word)

Represents a change marked with a revision mark. The **Revision** object is a member of the **[Revisions](Word.revisions.md)** collection. The **Revisions** collection includes all the revision marks in a range or document.


## Remarks

Use  **Revisions** (Index), where Index is the index number, to return a single **Revision** object. The index number represents the position of the revision in the range or document. The following example displays the author name for the first revision in section one of the active document.


```vb
MsgBox ActiveDocument.Sections(1).Range.Revisions(1).Author
```

The **Add** method isn't available for the **Revisions** collection. **Revision** objects are added when change tracking is enabled. Set the **TrackRevisions** property to **True** to track revisions made to the document text. The following example enables revision tracking and then inserts "Action " before the selection.




```vb
ActiveDocument.TrackRevisions = True 
Selection.InsertBefore "Action "
```


## Methods



|Name|
|:-----|
|[Accept](Word.Revision.Accept.md)|
|[Reject](Word.Revision.Reject.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Revision.Application.md)|
|[Author](Word.Revision.Author.md)|
|[Cells](Word.Revision.Cells.md)|
|[Creator](Word.Revision.Creator.md)|
|[Date](Word.Revision.Date.md)|
|[FormatDescription](Word.Revision.FormatDescription.md)|
|[Index](Word.Revision.Index.md)|
|[MovedRange](Word.Revision.MovedRange.md)|
|[Parent](Word.Revision.Parent.md)|
|[Range](Word.Revision.Range.md)|
|[Style](Word.Revision.Style.md)|
|[Type](Word.Revision.Type.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]