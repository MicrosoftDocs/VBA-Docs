---
title: WorksheetFunction.Substitute method (Excel)
keywords: vbaxl10.chm137128
f1_keywords:
- vbaxl10.chm137128
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Substitute
ms.assetid: 1e02eb86-6902-0073-33ea-8d9f08b4eb14
ms.date: 05/25/2019
localization_priority: Normal
---


# WorksheetFunction.Substitute method (Excel)

Substitutes new_text for old_text in a text string. Use **Substitute** when you want to replace specific text in a text string; use **Replace** when you want to replace any text that occurs in a specific location in a text string.


## Syntax

_expression_.**Substitute** (_Arg1_, _Arg2_, _Arg3_, _Arg4_)

_expression_ A variable that represents a **[WorksheetFunction](Excel.WorksheetFunction.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text - the text or the reference to a cell containing text for which you want to substitute characters.|
| _Arg2_|Required| **String**|Old_text - the text that you want to replace.|
| _Arg3_|Required| **String**|New_text - the text that you want to replace old_text with.|
| _Arg4_|Optional| **Variant**|Instance_num - specifies which occurrence of old_text you want to replace with new_text. If you specify instance_num, only that instance of old_text is replaced. Otherwise, every occurrence of old_text in text is changed to new_text.|

## Return value

**String**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
