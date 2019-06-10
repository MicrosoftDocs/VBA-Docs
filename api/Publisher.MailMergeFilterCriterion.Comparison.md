---
title: MailMergeFilterCriterion.Comparison property (Publisher)
keywords: vbapb10.chm6815748
f1_keywords:
- vbapb10.chm6815748
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion.Comparison
ms.assetid: ba815a39-35d6-803e-39c4-deba30646e66
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeFilterCriterion.Comparison property (Publisher)

Returns or sets an **[MsoFilterComparison](Office.MsoFilterComparison.md)** constant that represents how to compare the **[Cell.Column](Publisher.Cell.Column.md)** and **[CompareTo](Publisher.MailMergeFilterCriterion.CompareTo.md)** properties. Read/write.


## Syntax

_expression_.**Comparison**

_expression_ A variable that represents a **[MailMergeFilterCriterion](Publisher.MailMergeFilterCriterion.md)** object.


## Return value

MsoFilterComparison


## Remarks

The **Comparison** property value can be one of the **MsoFilterComparison** constants declared in the Microsoft Office type library.


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to WA. This example assumes that a mail merge data source is attached to the active publication.

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
 Next intItem 
 End With 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]