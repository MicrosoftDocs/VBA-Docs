---
title: CoAuthoring.Updates property (Word)
keywords: vbawd10.chm254869510
f1_keywords:
- vbawd10.chm254869510
ms.prod: word
api_name:
- Word.CoAuthoring.Updates
ms.assetid: 89c99cbd-1b97-24b1-f614-d7ade4f383bc
ms.date: 06/08/2017
localization_priority: Normal
---


# CoAuthoring.Updates property (Word)

Returns a  **[CoAuthUpdates](overview/Word.md)** collection that represents the most recent updates that were merged into the document. Read-only.


## Syntax

_expression_. `Updates`

 _expression_ An expression that returns a '[CoAuthoring](Word.CoAuthoring.md)' object.


## Example

The following code example gets the most recent updates that have been merged into the active document.


```vb
Dim allUpdates As CoAuthUpdates 
 
Set allUpdates = ActiveDocument.CoAuthoring.Updates
```


## See also


[CoAuthoring Object](Word.CoAuthoring.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]