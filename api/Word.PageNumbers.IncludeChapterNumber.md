---
title: PageNumbers.IncludeChapterNumber property (Word)
keywords: vbawd10.chm159776771
f1_keywords:
- vbawd10.chm159776771
ms.prod: word
api_name:
- Word.PageNumbers.IncludeChapterNumber
ms.assetid: 0ab2dbb4-4bf3-f878-1fd9-abca20dd790a
ms.date: 06/08/2017
localization_priority: Normal
---


# PageNumbers.IncludeChapterNumber property (Word)

 **True** if a chapter number is included with page numbers or a caption label. Read/write **Boolean**.


## Syntax

_expression_. `IncludeChapterNumber`

_expression_ A variable that represents a '[PageNumbers](Word.pagenumbers.md)' object.


## Example

This example adds page numbers in the footer for section one in the active document. The page numbers include the chapter number.


```vb
With ActiveDocument.Sections(1).Footers _ 
 (wdHeaderFooterPrimary).PageNumbers 
 .Add 
 .IncludeChapterNumber = True 
 .HeadingLevelForChapter = 1 
End With
```

This example adds the chapter number from the Heading 2 style to figure captions, sets the caption numbering style, and then inserts a new figure caption. The document should already contain a Heading 2 style with numbering.




```vb
With CaptionLabels(wdCaptionFigure) 
 .IncludeChapterNumber = True 
 .ChapterStyleLevel = 2 
 .NumberStyle = wdCaptionNumberStyleUppercaseLetter 
End With 
Selection.InsertCaption Label:="Figure", Title:=": History"
```


## See also


[PageNumbers Collection Object](Word.pagenumbers.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]