---
title: WorksheetFunction.Phonetic method (Excel)
keywords: vbaxl10.chm137248
f1_keywords:
- vbaxl10.chm137248
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Phonetic
ms.assetid: a1da7aa0-f913-e64b-8863-212f8a4e261d
ms.date: 05/24/2019
localization_priority: Normal
---


# WorksheetFunction.Phonetic method (Excel)

Extracts the phonetic (furigana) characters from a text string.


## Syntax

_expression_.**Phonetic** (_Arg1_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Range**|Reference - a text string or a reference to a single cell or a range of cells that contain a furigana text string.|

## Return value

**Double**


## Remarks

If reference is a range of cells, the furigana text string in the upper-left corner cell of the range is returned.
    
If the reference is a range of nonadjacent cells, the #N/A error value is returned. 
    


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]