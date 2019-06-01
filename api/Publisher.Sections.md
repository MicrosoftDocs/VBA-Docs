---
title: Sections object (Publisher)
keywords: vbapb10.chm7405567
f1_keywords:
- vbapb10.chm7405567
ms.prod: publisher
api_name:
- Publisher.Sections
ms.assetid: 429c03b8-b574-86db-c39d-551a4c753b04
ms.date: 06/01/2019
localization_priority: Normal
---


# Sections object (Publisher)

A collection of all the **[Section](Publisher.Section.md)** objects in the document.
 
## Remarks

Use **Item** (_index_), where (_index_) is the index number, to return a single **Section** object. Using **Sections** (_index_), where (_index_) is the index number, will also return a single **Section** object. 

Use **Count** to return the number of sections in the publication. 

Use **Add** (_StartPageIndex_), where _StartPageIndex_ is the index number of the page, to return a new section added to a document. A "Permission denied" error is returned if the page already contains a section head. 

Use **[Section.Delete](publisher.section.delete.md)** (_index_), where (_index_) is the index number, to delete the specified section from the document. A "Permission denied" error is returned if an attempt is made to delete the first section. 



## Example

The following example sets the number format and the starting number for the first section of the active document.

```vb
With ActiveDocument.Sections.Item(1) 
 .PageNumberFormat = pbPageNumberFormatArabic 
 .PageNumberStart = 1 
End With
```

<br/>

The following example continues the numbering from the previous section for the second section in the active document.

```vb
ActiveDocument.Sections(2).ContinueNumbersFromPreviousSection=True
```

<br/>

The following example displays the number of sections in the first open document.

```vb
MsgBox Documents(1).Sections.Count
```

<br/>

The following example adds a new section to the second page of the active document.

```vb
Dim objSection As Section 
Set objSection = ActiveDocument.Sections.Add(StartPageIndex:=2)
```

<br/>

The following example deletes all the sections of the active document except the first one.

> [!NOTE] 
> The iteration is from the last to the first to avoid a "Subscript out of range" error when accessing a deleted section in the **Sections** collection.
 
```vb
Dim i As Long 
For i = ActiveDocument.Sections.Count To 1 Step -1 
 If i = 1 Then Exit For 
 ActiveDocument.Sections(i).Delete 
Next i
```


## Methods

- [Add](Publisher.Sections.Add.md)

## Properties

- [Application](Publisher.Sections.Application.md)
- [Count](Publisher.Sections.Count.md)
- [Item](Publisher.Sections.Item.md)
- [Parent](Publisher.Sections.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]