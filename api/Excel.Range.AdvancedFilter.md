---
title: Range.AdvancedFilter method (Excel)
keywords: vbaxl10.chm144078
f1_keywords:
- vbaxl10.chm144078
api_name:
- Excel.Range.AdvancedFilter
ms.assetid: fe1a19fc-ab0f-6149-25d9-6102d5789757
ms.date: 05/10/2019
ms.localizationpriority: medium
---


# Range.AdvancedFilter method (Excel)

Filters or copies data from a list based on a criteria range. If the initial selection is a single cell, that cell's current region is used.


## Syntax

_expression_.**AdvancedFilter** (_Action_, _CriteriaRange_, _CopyToRange_, _Unique_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Action_|Required| **[XlFilterAction](Excel.XlFilterAction.md)**|One of the constants of **XlFilterAction** specifying whether to make a copy or filter the list in place.|
| _CriteriaRange_|Optional| **Variant**|The criteria range. If this argument is omitted, there are no criteria.|
| _CopyToRange_|Optional| **Variant**|The destination range for the copied rows if _Action_ is **xlFilterCopy**. Otherwise, this argument is ignored.|
| _Unique_|Optional| **Variant**| **True** to filter unique records only. **False** to filter all records that meet the criteria. The default value is **False**.|

## Return value

Variant


## Example

This example filters a database named Database based on a criteria range named Criteria.

```vb
Range("Database").AdvancedFilter _ 
 Action:=xlFilterInPlace, _ 
 CriteriaRange:=Range("Criteria")
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
