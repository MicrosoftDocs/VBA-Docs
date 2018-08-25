---
title: Errors.Item Property (Excel)
keywords: vbaxl10.chm700073
f1_keywords:
- vbaxl10.chm700073
ms.prod: excel
api_name:
- Excel.Errors.Item
ms.assetid: e7182924-48cb-d97d-93b4-b4f53542013e
ms.date: 06/08/2017
---


# Errors.Item Property (Excel)

Returns a single member of the  **[Error](Excel.Error.md)** object.


## Syntax

 _expression_. `Item`( `_Index_` )

 _expression_ A variable that represents an [Errors](./Excel.Errors.md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The index of the member.|

## Remarks

 _Index_ can also be one the following constants.

| **Constant** | **Description** |
|:----|:----|
| **xlEvaluateToError** | The cell evaluates to an error value.|
| **xlTextDate** | The cell contains a text date with 2 digit years.|
| **xlNumberAsText** | The cell contains a number stored as text.|
| **xlInconsistentFormula** | The cell contains an inconsistent formula for a region.|
| **xlOmittedCells** | The cell contains a formula omitting a cell for a region.|
| **xlUnlockedFormulaCells** | The cell which is unlocked contains a formula.|
| **xlEmptyCellReferences** | The cell contains a formula referring to empty cells.|

## See also


[Errors Object](Excel.Errors.md)

