---
title: TableOfContents.UseHyperlinks property (Word)
keywords: vbawd10.chm152240139
f1_keywords:
- vbawd10.chm152240139
ms.prod: word
api_name:
- Word.TableOfContents.UseHyperlinks
ms.assetid: 2ff74d58-6411-eb10-1ce4-86d0b8e37490
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfContents.UseHyperlinks property (Word)

Returns or sets whether entries in a table of contents should be formatted as hyperlinks when publishing to the Web. Read/write  **Boolean**.


## Syntax

_expression_. `UseHyperlinks`

_expression_ Required. A variable that represents a '[TableOfContents](Word.TableOfContents.md)' collection.


## Example

This example formats the first table of contents in the document using hyperlinks.


```vb
ActiveDocument.TableOfContents(1).UseHyperlinks = True
```


## See also


[TableOfContents Object](Word.TableOfContents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]