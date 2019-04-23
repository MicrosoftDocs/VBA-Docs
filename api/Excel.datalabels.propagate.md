---
title: DataLabels.Propagate method (Excel)
keywords: vbaxl10.chm584110
f1_keywords:
- vbaxl10.chm584110
ms.prod: excel
ms.assetid: cf81fe7c-fb9c-bcd5-bd29-aef898c9c265
ms.date: 04/23/2019
localization_priority: Normal
---


# DataLabels.Propagate method (Excel)

Enables you to take the contents and formatting of a single data label and apply it to every other data label in the series.


## Syntax

_expression_.**Propagate** (_Index_)

_expression_ A variable that represents a **[DataLabels](Excel.DataLabels(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|The index number of the data label to propagate.|

## Remarks

If the source data label supports fields that are incompatible with the destination data label, those fields will be inserted as their [FIELD NAME]. For example, if a data label on a pie series with a percentage field is propagated to a data label on a column chart, that percentage field will become an unresolved field showing [PERCENTAGE].

> [!NOTE] 
> Passing an argument of zero (0) resets the data labels to your current prototype.


## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]