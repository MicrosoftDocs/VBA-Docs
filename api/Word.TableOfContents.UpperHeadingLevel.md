---
title: TableOfContents.UpperHeadingLevel property (Word)
keywords: vbawd10.chm152240131
f1_keywords:
- vbawd10.chm152240131
ms.prod: word
api_name:
- Word.TableOfContents.UpperHeadingLevel
ms.assetid: 3b360b6b-a422-4af5-9121-200105b0ad19
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.UpperHeadingLevel property (Word)

Returns or sets the starting heading level for a table of contents. Read/write  **Long**.


## Syntax

_expression_. `UpperHeadingLevel`

_expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Remarks

This property corresponds to the starting value used with the \o switch for a Table of Contents (TOC) field. Use the  **[LowerHeadingLevel](Word.TableOfContents.LowerHeadingLevel.md)** property to set the ending heading level. For example, to set the TOC field syntax {TOC \o "1-3"}, set the **LowerHeadingLevel** property to 3 and the **UpperHeadingLevel** property to 1.


## Example

This example formats the first table of contents in the active document to compile all headings that are formatted with either the Heading 2 or Heading 3 style.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .UseHeadingStyles = True 
 .UseFields = False 
 .UpperHeadingLevel = 2 
 .LowerHeadingLevel = 3 
 End With 
End If
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]