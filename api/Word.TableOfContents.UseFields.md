---
title: TableOfContents.UseFields property (Word)
keywords: vbawd10.chm152240130
f1_keywords:
- vbawd10.chm152240130
ms.prod: word
api_name:
- Word.TableOfContents.UseFields
ms.assetid: 36d01961-ba9a-fe8d-d791-f892bea8b994
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.UseFields property (Word)

 **True** if Table of Contents Entry (TC) fields are used to create a table of contents or a table of figures. Read/write **Boolean**.


## Syntax

 _expression_. `UseFields`

 _expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example formats the first table of contents in the active document to use heading styles instead of TC fields.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .UseFields = False 
 .UseHeadingStyles = True 
 End With 
End If
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]