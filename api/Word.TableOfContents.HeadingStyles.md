---
title: TableOfContents.HeadingStyles property (Word)
keywords: vbawd10.chm152240134
f1_keywords:
- vbawd10.chm152240134
ms.prod: word
api_name:
- Word.TableOfContents.HeadingStyles
ms.assetid: 05cf7783-6b5d-bfbb-a417-1ae12d13f78e
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.HeadingStyles property (Word)

Returns a  **[HeadingStyles](Word.headingstyles.md)** object that represents additional styles used to compile a table of contents or table of figures (styles other than the Heading 1 - Heading 9 styles). Read-only.


## Syntax

_expression_. `HeadingStyles`

_expression_ A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example adds a style to the HeadingStyles collection and then displays the names of all the style in the collection.


```vb
Dim hsLoop As HeadingStyle 
 
If ActiveDocument.TablesOfContents.Count >=1 Then 
 ActiveDocument.TablesOfContents(1).HeadingStyles.Add _ 
 Style:="Title", Level:=2 
 For Each hsLoop In _ 
 ActiveDocument.TablesOfContents(1).HeadingStyles 
 MsgBox hsLoop.Style 
 Next hsLoop 
End If
```

This example adds a style named "Blue" to the HeadingStyles collection in a table of contents for Sales.doc.




```vb
With Documents("Sales.doc") 
 .Styles.Add Name:="Blue" 
 .TablesOfContents(1).UseHeadingStyles = True 
 .TablesOfContents(1).HeadingStyles.Add _ 
 Style:="Blue", Level:=4 
End With
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]