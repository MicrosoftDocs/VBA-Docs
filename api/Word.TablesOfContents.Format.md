---
title: TablesOfContents.Format property (Word)
keywords: vbawd10.chm152305666
f1_keywords:
- vbawd10.chm152305666
ms.prod: word
api_name:
- Word.TablesOfContents.Format
ms.assetid: ea94f93f-3fce-2b21-1f8b-675d5d3de96e
ms.date: 06/08/2017
localization_priority: Normal
---


# TablesOfContents.Format property (Word)

Returns or sets the formatting for the tables of contents in the specified document. Read/write  **WdTocFormat**.


## Syntax

 _expression_. `Format`

 _expression_ Required. A variable that represents a '[TablesOfContents](Word.tablesofcontents.md)' collection.


## Example

This example applies Classic formatting to the tables of contents in Report.doc.


```vb
Documents("Report.doc").TablesOfContents.Format = wdTOCClassic
```


## See also


[TablesOfContents Collection Object](Word.tablesofcontents.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]