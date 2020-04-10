---
title: Range.TCSCConverter method (Word)
keywords: vbawd10.chm157155827
f1_keywords:
- vbawd10.chm157155827
ms.prod: word
api_name:
- Word.Range.TCSCConverter
ms.assetid: 71684cdd-fca8-37b7-04fe-eeeb35dcfe66
ms.date: 06/08/2017
localization_priority: Normal
---


# Range.TCSCConverter method (Word)

Converts the specified range from Traditional Chinese to Simplified Chinese or vice versa.


## Syntax

_expression_. `TCSCConverter`( `_WdTCSCConverterDirection_` , `_CommonTerms_` , `_UseVariants_` )

_expression_ Required. A variable that represents a **[Range](Word.Range.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _WdTCSCConverterDirection_|Optional| **WdTCSCConverterDirection**|Specifies the direction in which text is converted. If omitted, the default value is **wdTCSCConverterDirectionAuto**, which converts in the appropriate direction based on the detected language of the specified range.|
| _UseVariants_|Optional| **Boolean**| **True** if Word uses Taiwan, Hong Kong SAR, and Macao SAR character variants. Can only be used if translating from Simplified Chinese to Traditional Chinese.|

## Example

This example converts the current selection from Simplified Chinese to Traditional Chinese. It converts common expressions intact and uses regional character variants.


```vb
Selection.Range.TCSCConverter _ 
 wdTCSCConverterDirectionSCTC, True, True
```


## See also


[Range Object](Word.Range.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]