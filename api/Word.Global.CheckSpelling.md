---
title: Global.CheckSpelling method (Word)
keywords: vbawd10.chm163119428
f1_keywords:
- vbawd10.chm163119428
ms.prod: word
api_name:
- Word.Global.CheckSpelling
ms.assetid: eb264c56-090f-b1af-3d0b-5ee5190345b7
ms.date: 06/08/2017
localization_priority: Normal
---


# Global.CheckSpelling method (Word)

Checks a string for spelling errors. Returns a  **Boolean** to indicate whether the string contains spelling errors. **True** if the string has no spelling errors.


## Syntax

_expression_.**CheckSpelling** (_Word_, _CustomDictionary_, _IgnoreUppercase_, _MainDictionary_, _CustomDictionary2_, _CustomDictionary3_, _CustomDictionary4_, _CustomDictionary5_, _CustomDictionary6_, _CustomDictionary7_, _CustomDictionary8_, _CustomDictionary9_, _CustomDictionary10_)

_expression_ Required. A variable that represents a '[Global](Word.Global.md)' object.


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


[Global Object](Word.Global.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]