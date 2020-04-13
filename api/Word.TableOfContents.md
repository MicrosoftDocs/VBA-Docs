---
title: TableOfContents object (Word)
ms.prod: word
api_name:
- Word.TableOfContents
ms.assetid: 629a03c1-ae97-649d-7ec4-25210b4b9ecd
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents object (Word)

Represents a single table of contents in a document. The **TableOfContents** object is a member of the **[TablesOfContents](Word.tablesofcontents.md)** collection. The **TablesOfContents** collection includes all the tables of contents in a document.


## Remarks

Use  **TablesOfContents** (Index), where Index is the index number, to return a single **TableOfContents** object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document.


```vb
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```

Use the **Add** method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfContents.Add Range:=myRange, _ 
 UseFields:=False, UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, _ 
 UpperHeadingLevel:=1
```


## Methods



|Name|
|:-----|
|[Delete](Word.TableOfContents.Delete.md)|
|[Update](Word.TableOfContents.Update.md)|
|[UpdatePageNumbers](Word.TableOfContents.UpdatePageNumbers.md)|

## Properties



|Name|
|:-----|
|[Application](Word.TableOfContents.Application.md)|
|[Creator](Word.TableOfContents.Creator.md)|
|[HeadingStyles](Word.TableOfContents.HeadingStyles.md)|
|[HidePageNumbersInWeb](Word.TableOfContents.HidePageNumbersInWeb.md)|
|[IncludePageNumbers](Word.TableOfContents.IncludePageNumbers.md)|
|[LowerHeadingLevel](Word.TableOfContents.LowerHeadingLevel.md)|
|[Parent](Word.TableOfContents.Parent.md)|
|[Range](Word.TableOfContents.Range.md)|
|[RightAlignPageNumbers](Word.TableOfContents.RightAlignPageNumbers.md)|
|[TabLeader](Word.TableOfContents.TabLeader.md)|
|[TableID](Word.TableOfContents.TableID.md)|
|[UpperHeadingLevel](Word.TableOfContents.UpperHeadingLevel.md)|
|[UseFields](Word.TableOfContents.UseFields.md)|
|[UseHeadingStyles](Word.TableOfContents.UseHeadingStyles.md)|
|[UseHyperlinks](Word.TableOfContents.UseHyperlinks.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]