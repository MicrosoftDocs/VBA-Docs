---
title: Application.CheckSpelling method (Word)
keywords: vbawd10.chm158335300
f1_keywords:
- vbawd10.chm158335300
ms.prod: word
api_name:
- Word.Application.CheckSpelling
ms.assetid: 88ea2134-cdbf-2bd5-bd6a-ff0c32a0f568
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.CheckSpelling method (Word)

Checks a string for spelling errors. Returns a  **Boolean** to indicate whether the string contains spelling errors. **True** if the string has no spelling errors.


## Syntax

_expression_.**CheckSpelling** (_Word_, _CustomDictionary_, _IgnoreUppercase_, _MainDictionary_, _CustomDictionary2_, _CustomDictionary3_, _CustomDictionary4_, _CustomDictionary5_, _CustomDictionary6_, _CustomDictionary7_, _CustomDictionary8_, _CustomDictionary9_, _CustomDictionary10_)

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**|The text whose spelling is to be checked.|
| _CustomDictionary_|Optional| **Variant**| Either an expression that returns a Dictionary object or the file name of the custom dictionary.|
| _IgnoreUppercase_|Optional| **Variant**| **True** if capitalization is ignored. If this argument is omitted, the current value of the **IgnoreUppercase** property is used.|
| _MainDictionary_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of the main dictionary.|
| _CustomDictionary2_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary3_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary4_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary5_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary6_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary7_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary8_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary9_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|
| _CustomDictionary10_|Optional| **Variant**|Either an expression that returns a  **Dictionary** object or the file name of an additional custom dictionary.|

## Return value

Boolean


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]