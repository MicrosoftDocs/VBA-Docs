---
title: TablesOfContents object (Word)
keywords: vbawd10.chm2324
f1_keywords:
- vbawd10.chm2324
ms.prod: word
ms.assetid: d0d0e5fc-e443-31ae-e1a9-15b945f1e318
ms.date: 06/08/2017
localization_priority: Normal
---


# TablesOfContents object (Word)

A collection of  **[TableOfContents](Word.TableOfContents.md)** objects that represent the tables of contents in a document.


## Remarks

Use the **TablesOfContents** property to return the **TablesOfContents** collection. The following example inserts a table of contents entry that references the selected text in the active document.


```vb
ActiveDocument.TablesOfContents.MarkEntry Range:=Selection.Range, _ 
 Level:=2, Entry:="Introduction"
```

Use the **Add** method to add a table of contents to a document. The following example adds a table of contents at the beginning of the active document. The example builds the table of contents from all paragraphs styled as either Heading 1, Heading 2, or Heading 3.




```vb
Set myRange = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfContents.Add Range:=myRange, _ 
 UseFields:=False, UseHeadingStyles:=True, _ 
 LowerHeadingLevel:=3, _ 
 UpperHeadingLevel:=1
```

Use  **TablesOfContents** (Index), where Index is the index number, to return a single **TableOfContents** object. The index number represents the position of the table of contents in the document. The following example updates the page numbers of the items in the first table of figures in the active document.




```vb
ActiveDocument.TablesOfContents(1).UpdatePageNumbers
```


## Methods



|Name|
|:-----|
|[Add](Word.TablesOfContents.Add.md)|
|[Item](Word.TablesOfContents.Item.md)|
|[MarkEntry](Word.TablesOfContents.MarkEntry.md)|

## Properties



|Name|
|:-----|
|[Application](Word.TablesOfContents.Application.md)|
|[Count](Word.TablesOfContents.Count.md)|
|[Creator](Word.TablesOfContents.Creator.md)|
|[Format](Word.TablesOfContents.Format.md)|
|[Parent](Word.TablesOfContents.Parent.md)|

## See also


[Word Object Model Reference](overview/Word/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
