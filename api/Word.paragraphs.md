---
title: Paragraphs object (Word)
ms.prod: word
ms.assetid: bdc7a183-2a98-7d47-c86a-5cecd6c91449
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraphs object (Word)

A collection of  **[Paragraph](Word.Paragraph.md)** objects in a selection, range, or document.


## Remarks

Use the  **Paragraphs** property to return the **Paragraphs** collection. The following example formats the selected paragraphs to be double-spaced and right-aligned.


```vb
With Selection.Paragraphs 
 .Alignment = wdAlignParagraphRight 
 .LineSpacingRule = wdLineSpaceDouble 
End With
```

Use the  **Add**, **InsertParagraph**, **InsertParagraphAfter**, or **InsertParagraphBefore** method to add a new paragraph to a document. The following example adds a new paragraph before the first paragraph in the selection.




```vb
Selection.Paragraphs.Add Range:=Selection.Paragraphs(1).Range
```

The following example also adds a paragraph before the first paragraph in the selection.




```vb
Selection.Paragraphs(1).Range.InsertParagraphBefore
```

Use  **Paragraphs** (Index), where Index is the index number, to return a single **Paragraph** object. The following example right aligns the first paragraph in the active document.




```vb
ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphRight
```

The  **Count** property for this collection in a document returns the number of items in the main story only. To count items in other stories use the collection with the **Range** object.


## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
