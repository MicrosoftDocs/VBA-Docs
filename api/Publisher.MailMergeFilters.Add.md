---
title: MailMergeFilters.Add method (Publisher)
keywords: vbapb10.chm6750212
f1_keywords:
- vbapb10.chm6750212
ms.prod: publisher
api_name:
- Publisher.MailMergeFilters.Add
ms.assetid: ab114dda-d144-7c5f-88b0-930cadcf53db
ms.date: 06/11/2019
localization_priority: Normal
---


# MailMergeFilters.Add method (Publisher)

Adds a new filter criterion to the specified **MailMergeFilters** object.


## Syntax

_expression_.**Add** (_Column_, _Comparison_, _Conjunction_, _bstrCompareTo_, _DeferUpdate_)

_expression_ A variable that represents a **[MailMergeFilters](Publisher.MailMergeFilters.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_Column_|Required| **String**|The name of the table in the data source.|
|_Comparison_ |Required| **[MsoFilterComparison](Office.MsoFilterComparison.md)**  |How the data in the table is filtered. Can be one of the **MsoFilterComparison** constants.|
|_Conjunction_|Required| **[MsoFilterConjunction](office.msofilterconjunction.md)**|How this filter relates to other filters in the **MailMergeFilters** object. Can be one of the **MsoFilterConjunction** constants.|
|_bstrCompareTo_|Optional| **String**|If the _Comparison_ argument is something other than **msoFilterComparisonIsBlank** or **msoFilterComparisonIsNotBlank**, _bstrCompareTo_ is a string to which the data in the table is compared.|
|_DeferUpdate_|Optional| **Boolean**| **True** to queue the filters and apply them when the **ApplyFilter** method is called. **False** to apply the filter condition immediately. Default is **False**.|



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]