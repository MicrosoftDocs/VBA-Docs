---
title: TableOfFigures.UseHyperlinks property (Word)
keywords: vbawd10.chm153157645
f1_keywords:
- vbawd10.chm153157645
ms.prod: word
api_name:
- Word.TableOfFigures.UseHyperlinks
ms.assetid: 63568e7b-b3ac-6fda-e0a3-48eb46c2f48d
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfFigures.UseHyperlinks property (Word)

Returns or sets whether entries in a table of figures should be formatted as hyperlinks when publishing to the Web. Read/write  **Boolean**.


## Syntax

 _expression_. `UseHyperlinks`

 _expression_ Required. A variable that represents a '[TableOfFigures](Word.TableOfFigures.md)' collection.


## Example

This example formats the first table of figures in the document using hyperlinks.


```vb
ActiveDocument.TableOfFigures(1).UseHyperlinks = True
```


## See also


[TableOfFigures Object](Word.TableOfFigures.md)

