---
title: TableOfAuthorities.TabLeader property (Word)
keywords: vbawd10.chm152109068
f1_keywords:
- vbawd10.chm152109068
ms.prod: word
api_name:
- Word.TableOfAuthorities.TabLeader
ms.assetid: b437456d-30a2-8e47-2527-dab0b6a43755
ms.date: 06/08/2017
localization_priority: Normal
---


# TableOfAuthorities.TabLeader property (Word)

Returns or sets the leader character that appears between entries and their associated page numbers in a table of authorities. Read/write  **WdTabLeader**.


## Syntax

_expression_. `TabLeader`

_expression_ Required. A variable that represents a '[TableOfAuthorities](Word.TableOfAuthorities.md)' collection.


## Example

This example formats the tables of authorities in Sales.doc to use a dotted tab leader.


```vb
For Each aTOA In Documents("Sales.doc").TablesOfAuthorities 
 aTOA.TabLeader = wdTabLeaderDots 
Next aTOA
```


## See also


[TableOfAuthorities Object](Word.TableOfAuthorities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]