---
title: TableOfAuthorities.IncludeCategoryHeader property (Word)
keywords: vbawd10.chm152109065
f1_keywords:
- vbawd10.chm152109065
ms.prod: word
api_name:
- Word.TableOfAuthorities.IncludeCategoryHeader
ms.assetid: 63118a82-28ac-f5c9-790d-0a8ea4926858
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfAuthorities.IncludeCategoryHeader property (Word)

 **True** if the category name for a group of entries appears in the table of authorities. Read/write **Boolean**.


## Syntax

 _expression_. `IncludeCategoryHeader`

 _expression_ An expression that returns a '[TableOfAuthorities](Word.TableOfAuthorities.md)' object.


## Remarks

Corresponds to the \h switch for a Table of Authorities (TOA) field.


## Example

This example includes the category name for each table of authorities in the active document.


```vb
Dim toaLoop As TableOfAuthorities 
 
For Each toaLoop In ActiveDocument.TablesOfAuthorities 
 toaLoop.IncludeCategoryHeader = True 
Next toaLoop
```


## See also


[TableOfAuthorities Object](Word.TableOfAuthorities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]