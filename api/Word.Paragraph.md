---
title: Paragraph object (Word)
keywords: vbawd10.chm2391
f1_keywords:
- vbawd10.chm2391
ms.prod: word
api_name:
- Word.Paragraph
ms.assetid: 0a704079-a082-4ab1-841b-fc0d49dd26d4
ms.date: 06/08/2017
localization_priority: Normal
---


# Paragraph object (Word)

Represents a single paragraph in a selection, range, or document. The **Paragraph** object is a member of the **[Paragraphs](Word.paragraphs.md)** collection. The **Paragraphs** collection includes all the paragraphs in a selection, range, or document.


## Remarks

Use  **Paragraphs** (Index), where Index is the index number, to return a single **Paragraph** object. The following example right aligns the first paragraph in the active document.


```vb
ActiveDocument.Paragraphs(1).Alignment = wdAlignParagraphRight
```

Use the **Add**, **InsertParagraph**, **InsertParagraphAfter**, or **InsertParagraphBefore** method to add a new, blank paragraph to a document. The following example adds a paragraph mark before the first paragraph in the selection.




```vb
Selection.Paragraphs.Add Range:=Selection.Paragraphs(1).Range
```

The following example also adds a paragraph mark before the first paragraph in the selection.




```vb
Selection.Paragraphs(1).Range.InsertParagraphBefore
```


## See also



[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
