---
title: MailMergeDataSource.Close method (Publisher)
keywords: vbapb10.chm6291493
f1_keywords:
- vbapb10.chm6291493
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.Close
ms.assetid: c215743b-590a-6db9-e902-b9179b67bb8e
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.Close method (Publisher)

Closes the specified mail merge data source, cancels the mail merge, and converts all mail merge data fields to plain text.


## Syntax

_expression_.**Close**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Remarks

Closing a mail merge data source deletes the shape that represents the catalog merge area of the publication page associated with the data source.


## Example

The following example closes the data source for the active mail merge publication.

```vb
ActiveDocument.MailMerge.DataSource.Close
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]