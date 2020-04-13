---
title: Section object (Word)
keywords: vbawd10.chm2393
f1_keywords:
- vbawd10.chm2393
ms.prod: word
api_name:
- Word.Section
ms.assetid: 3fe563d8-fc05-c17a-e67b-c50eea7e7f13
ms.date: 06/08/2017
localization_priority: Normal
---


# Section object (Word)

Represents a single section in a selection, range, or document. The **Section** object is a member of the **[Sections](Word.sections.md)** collection. The **Sections** collection includes all the sections in a selection, range, or document.


## Remarks

Use  **Sections** (Index), where Index is the index number, to return a single **Section** object. The following example changes the left and right page margins for the first section in the active document.


```vb
With ActiveDocument.Sections(1).PageSetup 
 .LeftMargin = InchesToPoints(0.5) 
 .RightMargin = InchesToPoints(0.5) 
End With
```

Use the **Add** method or the **InsertBreak** method to add a new section to a document. The following example adds a new section at the beginning of the active document.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.Sections.Add Range:=myRange 
myRange.InsertParagraphAfter
```

The following example adds a section break above the first paragraph in the selection.




```vb
Selection.Paragraphs(1).Range.InsertBreak _ 
 Type:=wdSectionBreakContinuous
```


> [!NOTE] 
> The **Headers** and **Footers** properties of the specified **Section** object return a **HeadersFooters** object.


## Properties



|Name|
|:-----|
|[Application](Word.Section.Application.md)|
|[Borders](Word.Section.Borders.md)|
|[Creator](Word.Section.Creator.md)|
|[Footers](Word.Section.Footers.md)|
|[Headers](Word.Section.Headers.md)|
|[Index](Word.Section.Index.md)|
|[PageSetup](Word.Section.PageSetup.md)|
|[Parent](Word.Section.Parent.md)|
|[ProtectedForForms](Word.Section.ProtectedForForms.md)|
|[Range](Word.Section.Range.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
