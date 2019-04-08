---
title: TableOfContents.LowerHeadingLevel property (Word)
keywords: vbawd10.chm152240132
f1_keywords:
- vbawd10.chm152240132
ms.prod: word
api_name:
- Word.TableOfContents.LowerHeadingLevel
ms.assetid: 02bd1965-b3a1-e09a-fb08-62862e87536b
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.LowerHeadingLevel property (Word)

Returns or sets the ending heading level for a table of contents or table of figures. Read/write  **Long**.


## Syntax

_expression_. `LowerHeadingLevel`

_expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Remarks

This property corresponds to the ending value used with the \o switch for a Table of Contents (TOC) field. Use the  **[UpperHeadingLevel](Word.TableOfContents.UpperHeadingLevel.md)** property to set the starting heading level. For example, to set the TOC field syntax {TOC \o "1-3"}, set the **LowerHeadingLevel** property to 3 and the **UpperHeadingLevel** property to 1.


## Example

This example formats the first table of contents in the active document to show entries formatted with the Heading 2, Heading 3, and Heading 4 styles.


```vb
If ActiveDocument.TablesOfContents.Count >= 1 Then 
 With ActiveDocument.TablesOfContents(1) 
 .UseHeadingStyles = True 
 .UpperHeadingLevel = 2 
 .LowerHeadingLevel = 4 
 End With 
End If
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]