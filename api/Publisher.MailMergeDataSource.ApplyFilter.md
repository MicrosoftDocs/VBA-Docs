---
title: MailMergeDataSource.ApplyFilter method (Publisher)
keywords: vbapb10.chm6291492
f1_keywords:
- vbapb10.chm6291492
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.ApplyFilter
ms.assetid: a94af75c-e558-7160-76c9-c0f8c3fb317d
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.ApplyFilter method (Publisher)

Applies a filter to a mail merge data source to remove (or filter out) specified records containing (or not containing) specific data.


## Syntax

_expression_.**ApplyFilter**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Example

This example adds a new filter that removes all records with a blank Region field and then applies the filter to the active publication. This example assumes that a mail merge data source is attached to the active publication.

```vb
Sub FilterDataSource() 
 With ActiveDocument.MailMerge.DataSource 
 .Filters.Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]