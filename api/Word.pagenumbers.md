---
title: PageNumbers object (Word)
ms.prod: word
ms.assetid: 9090f96e-d898-ace6-35fa-f6e59c527ea2
ms.date: 06/08/2017
localization_priority: Normal
---


# PageNumbers object (Word)

A collection of  **PageNumber** objects that represent the page numbers in a single header or footer.


## Remarks

Use the **PageNumbers** property to return the **PageNumbers** collection. The following example starts page numbering at 3 for the first section in the active document.


```vb
ActiveDocument.Sections(1).Footers(wdHeaderFooterPrimary) _ 
 .PageNumbers.StartingNumber = 3
```

Use the **Add** method to add page numbers to a header or footer. The following example adds a page number to the primary footer in the first section.




```vb
With ActiveDocument.Sections(1) 
 .Footers(wdHeaderFooterPrimary).PageNumbers.Add _ 
 PageNumberAlignment:=wdAlignPageNumberLeft, _ 
 FirstPage:=False 
End With
```

To add or change page numbers in a document with multiple sections, modify the page numbers in each section or set the **LinkToPrevious** property to **True**.

Use  **PageNumbers** (_index_), where _index_ is the index number, to return a single **PageNumber** object. In most cases, a header or footer contains only one page number, which is index number 1. The following example centers the first page number in the primary header in the first section.




```vb
ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary) _ 
 .PageNumbers(1).Alignment = wdAlignPageNumberCenter
```


## Methods



|Name|
|:-----|
|[Add](Word.PageNumbers.Add.md)|
|[Item](Word.PageNumbers.Item.md)|

## Properties



|Name|
|:-----|
|[Application](Word.PageNumbers.Application.md)|
|[ChapterPageSeparator](Word.PageNumbers.ChapterPageSeparator.md)|
|[Count](Word.PageNumbers.Count.md)|
|[Creator](Word.PageNumbers.Creator.md)|
|[DoubleQuote](Word.PageNumbers.DoubleQuote.md)|
|[HeadingLevelForChapter](Word.PageNumbers.HeadingLevelForChapter.md)|
|[IncludeChapterNumber](Word.PageNumbers.IncludeChapterNumber.md)|
|[NumberStyle](Word.PageNumbers.NumberStyle.md)|
|[Parent](Word.PageNumbers.Parent.md)|
|[RestartNumberingAtSection](Word.PageNumbers.RestartNumberingAtSection.md)|
|[ShowFirstPageNumber](Word.PageNumbers.ShowFirstPageNumber.md)|
|[StartingNumber](Word.PageNumbers.StartingNumber.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
