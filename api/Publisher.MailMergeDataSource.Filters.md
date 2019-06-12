---
title: MailMergeDataSource.Filters property (Publisher)
keywords: vbapb10.chm6291463
f1_keywords:
- vbapb10.chm6291463
ms.prod: publisher
api_name:
- Publisher.MailMergeDataSource.Filters
ms.assetid: 7b8fa974-08e5-9691-c69d-314eb6a5c651
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeDataSource.Filters property (Publisher)

Returns a **[MailMergeFilters](Publisher.MailMergeFilters.md)** object that represents filters applied to the mail merge or catalog merge data source.


## Syntax

_expression_.**Filters**

_expression_ A variable that represents a **[MailMergeDataSource](Publisher.MailMergeDataSource.md)** object.


## Return value

MailMergeFilters


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