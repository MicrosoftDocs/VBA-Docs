---
title: Application.GetPhonetic method (Excel)
keywords: vbaxl10.chm133245
f1_keywords:
- vbaxl10.chm133245
ms.prod: excel
api_name:
- Excel.Application.GetPhonetic
ms.assetid: 530be07e-04ed-81c5-3b12-93b78e494a3b
ms.date: 04/04/2019
localization_priority: Normal
---


# Application.GetPhonetic method (Excel)

Returns the Japanese phonetic text of the specified text string. This method is available to you only if you have selected or installed Japanese language support for Microsoft Office.


## Syntax

_expression_.**GetPhonetic** (_Text_)

_expression_ A variable that represents an **[Application](Excel.Application(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Text_|Optional| **Variant**|Specifies the text to be converted to phonetic text. If you omit this argument, the next possible phonetic text string (if any) of the previously specified _Text_ is returned. If there are no more possible phonetic text strings, an empty string is returned.|

## Return value

String


## Example

This example displays all the possible phonetic text strings from the specified string.

```vb
strPhoText = Application.GetPhonetic("??") 
While strPhoText <> "" 
    MsgBox strPhoText 
    strPhoText = Application.GetPhonetic() 
Wend
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]