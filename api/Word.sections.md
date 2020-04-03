---
title: Sections object (Word)
keywords: vbawd10.chm2394
f1_keywords:
- vbawd10.chm2394
ms.prod: word
ms.assetid: cf6f77ba-9eee-5614-e697-bc031c4c6dcd
ms.date: 06/08/2017
localization_priority: Normal
---


# Sections object (Word)

A collection of  **Section** objects in a selection, range, or document.


## Remarks

Use the  **Sections** property to return the **Sections** collection. The following example inserts text at the end of the last section in the active document.


```vb
With ActiveDocument.Sections.Last.Range 
 .Collapse Direction:=wdCollapseEnd 
 .InsertAfter "end of document" 
End With
```

Use the  **Add** method or the **InsertBreak** method to add a new section to a document. The following example adds a new section at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Sections.Add Range:=myRange 
myRange.InsertParagraphAfter
```

The following example displays the number of sections in the active document, adds a section break above the first paragraph in the selection, and then displays the number of sections again.




```vb
MsgBox ActiveDocument.Sections.Count & " sections" 
Selection.Paragraphs(1).Range.InsertBreak _ 
 Type:=wdSectionBreakContinuous 
MsgBox ActiveDocument.Sections.Count & " sections"
```

Use  **Sections** (_index_), where _index_ is the index number, to return a single **Section** object. The following example changes the left and right page margins for the first section in the active document.




```vb
With ActiveDocument.Sections(1).PageSetup 
 .LeftMargin = InchesToPoints(0.5) 
 .RightMargin = InchesToPoints(0.5) 
End With
```


## Methods



|Name|
|:-----|
|[Add](Word.Sections.Add.md)|
|[Item](Word.Sections.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.Sections.Application.md)|
|[Count](Word.Sections.Count.md)|
|[Creator](Word.Sections.Creator.md)|
|[First](Word.Sections.First.md)|
|[Last](Word.Sections.Last.md)|
|[PageSetup](Word.Sections.PageSetup.md)|
|[Parent](Word.Sections.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
