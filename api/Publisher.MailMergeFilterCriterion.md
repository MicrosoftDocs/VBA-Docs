---
title: MailMergeFilterCriterion object (Publisher)
keywords: vbapb10.chm6881279
f1_keywords:
- vbapb10.chm6881279
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion
ms.assetid: 2814890f-009b-b277-3ea4-c1f167a5e1c9
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeFilterCriterion object (Publisher)

Represents a filter to be applied to an attached mail merge or catalog merge data source. The **MailMergeFilterCriterion** object is a member of the **[MailMergeFilters](Publisher.MailMergeFilters.md)** object.
 
## Remarks

Each filter is a line in a query string. Use the **Column**, **Comparison**, **CompareTo**, and **Conjunction** properties to return or set the data source query criterion. 

Use the **[Add](Publisher.MailMergeFilters.Add.md)** method of the **MailMergeFilters** object to add a new filter criterion to the query. 


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to WA. This example assumes that a data source is attached to the active publication.

```vb
Sub SetQueryCriterion() 
 Dim intItem As Integer 
 With ActiveDocument.MailMerge.DataSource.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next 
 End With 
End Sub
```

<br/>

This example adds a new line to the query string and then applies the combined filter to the data source. This example assumes that a data source is attached to the active publication.

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


## Properties

- [Application](Publisher.MailMergeFilterCriterion.Application.md)
- [Column](Publisher.MailMergeFilterCriterion.Column.md)
- [CompareTo](Publisher.MailMergeFilterCriterion.CompareTo.md)
- [Comparison](Publisher.MailMergeFilterCriterion.Comparison.md)
- [Conjunction](Publisher.MailMergeFilterCriterion.Conjunction.md)
- [Creator](Publisher.MailMergeFilterCriterion.Creator.md)
- [Index](Publisher.MailMergeFilterCriterion.Index.md)
- [Parent](Publisher.MailMergeFilterCriterion.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]