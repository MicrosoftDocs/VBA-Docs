---
title: MailMergeFilters object (Publisher)
keywords: vbapb10.chm6815743
f1_keywords:
- vbapb10.chm6815743
ms.prod: publisher
api_name:
- Publisher.MailMergeFilters
ms.assetid: 3a91c67f-6cc2-1d67-3382-04ead84f6f09
ms.date: 05/31/2019
localization_priority: Normal
---


# MailMergeFilters object (Publisher)

Represents all the filters to apply to the data source attached to the mail merge or catalog merge publication. The **MailMergeFilters** object is composed of **[MailMergeFilterCriterion](Publisher.MailMergeFilterCriterion.md)** objects.
 
## Remarks

Use the **Add** method to add a new filter criterion to the query. 

Use the **Item** method to access an individual filter criterion. 


## Example

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

<br/>

This example loops through all the filter criterion and if it finds one with a value of Region, changes it to remove from the mail merge all records that are not equal to WA. This example assumes that a data source is attached to the active publication.

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


## Methods

- [Add](Publisher.MailMergeFilters.Add.md)
- [Delete](Publisher.MailMergeFilters.Delete.md)
- [Item](Publisher.MailMergeFilters.Item.md)

## Properties

- [Application](Publisher.MailMergeFilters.Application.md)
- [Count](Publisher.MailMergeFilters.Count.md)
- [Creator](Publisher.MailMergeFilters.Creator.md)
- [Parent](Publisher.MailMergeFilters.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]