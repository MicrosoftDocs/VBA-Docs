---
title: MailMergeFilterCriterion.Conjunction Property (Publisher)
keywords: vbapb10.chm6815750
f1_keywords:
- vbapb10.chm6815750
ms.prod: publisher
api_name:
- Publisher.MailMergeFilterCriterion.Conjunction
ms.assetid: 79365a25-97fd-a18f-7815-eaccf4c5bdca
ms.date: 06/08/2017
localization_priority: Normal
---


# MailMergeFilterCriterion.Conjunction Property (Publisher)

Returns or sets an  **MsoFilterConjunction** constant that represents how a filter criterion relates to other filter criteria in the **[MailMergeFilters](Publisher.MailMergeFilters.md)** object. Read/write.


## Syntax

 _expression_. **Conjunction**

 _expression_ A variable that represents a  **MailMergeFilterCriterion** object.


## Return value

MsoFilterConjunction


## Remarks

The  **Conjunction** property value can be one of the following **MsoFilterConjunction** constants declared in the Microsoft Office type library.



| **msoFilterConjunctionAnd**|
| **msoFilterConjunctionOr**|

## Example

The following example changes an existing filter to remove from the mail merge all records that do not have a Region field equal to "WA", and then adds the filter to the following filter, so that the filter criteria must match both filters combined and not just one or the other.


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]