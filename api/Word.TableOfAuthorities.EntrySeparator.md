---
title: TableOfAuthorities.EntrySeparator property (Word)
keywords: vbawd10.chm152109063
f1_keywords:
- vbawd10.chm152109063
ms.prod: word
api_name:
- Word.TableOfAuthorities.EntrySeparator
ms.assetid: d063aabf-d50e-d66b-4450-5c589d79d3be
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfAuthorities.EntrySeparator property (Word)

Returns or sets the characters (up to five) that separate a table of authorities entry and its page number. Read/write  **String**.


## Syntax

_expression_. `EntrySeparator`

_expression_ A variable that represents a '[TableOfAuthorities](Word.TableOfAuthorities.md)' object.


## Remarks

The default is a tab character with a dotted leader. Corresponds to the \e switch for a TOA (Table of Authorities) field.


## Example

This example inserts a table of authorities into the active document and then formats the table to use a comma between the entries and their corresponding page numbers.


```vb
Dim rngTemp As Range 
Dim toaLoop As TableOfAuthorities 
 
Set rngTemp = ActiveDocument.Range(Start:=0, End:=0) 
ActiveDocument.TablesOfAuthorities.Add _ 
 Range:=rngTemp, Category:=1 
For Each toaLoop In ActiveDocument.TablesOfAuthorities 
 toaLoop.EntrySeparator = ", " 
Next toaLoop
```

This example returns the entry separator for the first table of authorities.




```vb
Dim strSeparator 
 
strSeparator = _ 
 ActiveDocument.TablesOfAuthorities(1).EntrySeparator
```


## See also


[TableOfAuthorities Object](Word.TableOfAuthorities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]