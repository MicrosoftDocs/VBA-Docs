---
title: MailMergeFilterCriterion.CompareTo Property (Publisher)
keywords: vbapb10.chm6815749
f1_keywords:
- vbapb10.chm6815749
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion.CompareTo
ms.assetid: 6e81fa38-a5d7-8421-6722-a18c5e9a8229
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFilterCriterion.CompareTo Property (Publisher)

Returns or sets a  **String** that represents the text to compare in the query filter criterion. Read/write.


## Syntax

 _expression_. **CompareTo**

 _expression_ A variable that represents a  **MailMergeFilterCriterion** object.


## Return value

String


## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA". This example assumes that a mail merge data source is attached to the active publication.


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