---
title: WorksheetFunction.ReplaceB method (Excel)
keywords: vbaxl10.chm137156
f1_keywords:
- vbaxl10.chm137156
ms.prod: excel
api_name:
- Excel.WorksheetFunction.ReplaceB
ms.assetid: 8853dcdd-6cd0-6ac5-1a71-27054f2a4776
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.ReplaceB method (Excel)

Replaces part of a text string, based on the number of bytes that you specify, with a different text string. 


## Syntax

_expression_.**ReplaceB** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Old_text - text in which you want to replace some characters.|
| _Arg2_|Required| **Double**|Start_num - the position of the character in old_text that you want to replace with new_text.|
| _Arg3_|Required| **Double**|Num_chars - the number of characters in old_text that you want **Replace** to replace with new_text.|
| _Arg4_|Required| **String**|New_text - the text that will replace characters in old_text.|

## Return value

**String**


## Remarks

**Replace** is intended for use with languages that use the single-byte character set (SBCS), whereas **ReplaceB** is intended for use with languages that use the double-byte character set (DBCS). The default language setting on your computer affects the return value in the following way:

- **Replace** always counts each character, whether single-byte or double-byte, as 1, no matter what the default language setting is.
    
- **ReplaceB** counts each double-byte character as 2 when you have enabled the editing of a language that supports DBCS, and then sets it as the default language. Otherwise, **ReplaceB** counts each character as 1.
    
- The languages that support DBCS include Japanese, Chinese (Simplified), Chinese (Traditional), and Korean. 



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]