---
title: ODSOFilters.Add method (Office)
keywords: vbaof11.chm241004
f1_keywords:
- vbaof11.chm241004
ms.prod: office
api_name:
- Office.ODSOFilters.Add
ms.assetid: ced18180-09bf-7663-66d5-7543ac7f6b84
ms.date: 01/22/2019
localization_priority: Normal
---


# ODSOFilters.Add method (Office)

Adds a new filter to the **ODSOFilters** collection.


## Syntax

_expression_.**Add** (_Column_, _Comparison_, _Conjunction_, _bstrCompareTo_, _DeferUpdate_)

_expression_ Required. A variable that represents an **[ODSOFilters](Office.ODSOFilters.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Column_|Required|**String**|The name of the table in the data source.|
| _Comparison_|Required|**[MsoFilterComparison](office.msofiltercomparison.md)**|How the data in the table is filtered.|
| _Conjunction_|Required|**[MsoFilterConjunction](office.msofilterconjunction.md)**|Determines how this filter relates to other filters in the **ODSOFilters** object.|
| _bstrCompareTo_|Optional|**String**|If the _Comparison_ argument is something other than **msoFilterComparisonIsBlank** or **msoFilterComparisonIsNotBlank**, _bstrCompareTo_ is a string to which the data in the table is compared.|
| _DeferUpdate_|Optional|**Boolean**|Specifies whether to delay updating the filter. Default is **False**.|

## See also

- [ODSOFilters object members](overview/library-reference/odsofilters-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]