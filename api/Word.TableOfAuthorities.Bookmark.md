---
title: TableOfAuthorities.Bookmark property (Word)
keywords: vbawd10.chm152109060
f1_keywords:
- vbawd10.chm152109060
ms.prod: word
api_name:
- Word.TableOfAuthorities.Bookmark
ms.assetid: 72cc5292-882c-df16-1b3e-9ed182be7ce7
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfAuthorities.Bookmark property (Word)

Returns or sets the name of the bookmark from which to collect table of authorities entries. Read/write  **String**.


## Syntax

_expression_. `Bookmark`

_expression_ A variable that represents a '[TableOfAuthorities](Word.TableOfAuthorities.md)' object.


## Remarks

The **Bookmark** property corresponds to the \b switch for a TOA (Table of Authorities) field.


## Example

If a table of authorities exists in the active document, the entries are collected from the area defined by the bookmark named "area."


```vb
If ActiveDocument.TablesOfAuthorities.Count >= 1 Then 
 ActiveDocument.TablesOfAuthorities(1).Bookmark = "area" 
End If
```


## See also


[TableOfAuthorities Object](Word.TableOfAuthorities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]