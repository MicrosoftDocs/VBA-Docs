---
title: MailMergeFilterCriterion Object (Publisher)
keywords: vbapb10.chm6881279
f1_keywords:
- vbapb10.chm6881279
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion
ms.assetid: 2814890f-009b-b277-3ea4-c1f167a5e1c9
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFilterCriterion Object (Publisher)

Represents a filter to be applied to an attached mail merge or catalog merge data source. The  **MailMergeFilterCriterion** object is a member of the **MailMergeFilters** object.
 


## Example

Each filter is a line in a query string. Use the  **[Column](Publisher.MailMergeFilterCriterion.Column.md)**, **[Comparison](Publisher.MailMergeFilterCriterion.Comparison.md)**, **[CompareTo](Publisher.MailMergeFilterCriterion.CompareTo.md)**, and **[Conjunction](Publisher.MailMergeFilterCriterion.Conjunction.md)** properties to return or set the data source query criterion. The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA". This example assumes that a data source is attached to the active publication.
 

 

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

Use the  **[Add](Publisher.MailMergeFilters.Add.md)** method of the **MailMergeFilters** object to add a new filter criterion to the query. This example adds a new line to the query string and then applies the combined filter to the data source. This example assumes that a data source is attached to the active publication.
 

 



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



|Name|
|:-----|
|[Application](Publisher.MailMergeFilterCriterion.Application.md)|
|[Column](Publisher.MailMergeFilterCriterion.Column.md)|
|[CompareTo](Publisher.MailMergeFilterCriterion.CompareTo.md)|
|[Comparison](Publisher.MailMergeFilterCriterion.Comparison.md)|
|[Conjunction](Publisher.MailMergeFilterCriterion.Conjunction.md)|
|[Creator](Publisher.MailMergeFilterCriterion.Creator.md)|
|[Index](Publisher.MailMergeFilterCriterion.Index.md)|
|[Parent](Publisher.MailMergeFilterCriterion.Parent.md)|

