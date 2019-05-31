---
title: Section object (Publisher)
keywords: vbapb10.chm7471103
f1_keywords:
- vbapb10.chm7471103
ms.prod: publisher
api_name:
- Publisher.Section
ms.assetid: 7e92a8de-ed66-564b-2657-cef0fc2392b8
ms.date: 06/01/2019
localization_priority: Normal
---


# Section object (Publisher)

Represents a section of a publication or document.
 
## Remarks

Use **[Sections.Item](publisher.sections.item.md)** (_index_), where _index_ is the index number, to return a single **Section** object. 

Use **[Sections.Add](publisher.sections.add.md)** (_StartPageIndex_), where _StartPageIndex_ is the index number of the page, to return a new section added to a document. A "Permission denied" error is returned if the page already contains a section head. 

## Example

The following example sets a **Section** object to the first section in the **Sections** collection of the active document.
 
```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Item(1)
```

<br/>

The following example adds a **Section** object to the second page of the active document.

```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```


## Methods

- [Delete](Publisher.Section.Delete.md)

## Properties

- [Application](Publisher.Section.Application.md)
- [ContinueNumbersFromPreviousSection](Publisher.Section.ContinueNumbersFromPreviousSection.md)
- [PageNumberFormat](Publisher.Section.PageNumberFormat.md)
- [PageNumberStart](Publisher.Section.PageNumberStart.md)
- [Parent](Publisher.Section.Parent.md)
- [ShowHeaderFooterOnFirstPage](Publisher.Section.ShowHeaderFooterOnFirstPage.md)
- [StartPageIndex](Publisher.Section.StartPageIndex.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]