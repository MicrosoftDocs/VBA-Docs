---
title: Range.CheckSpelling method (Excel)
keywords: vbaxl10.chm144093
f1_keywords:
- vbaxl10.chm144093
ms.prod: excel
api_name:
- Excel.Range.CheckSpelling
ms.assetid: 22540515-0b0f-ce3c-0145-e46d477d1b5f
ms.date: 05/10/2019
localization_priority: Normal
---


# Range.CheckSpelling method (Excel)

Checks the spelling of an object.


## Syntax

_expression_.**CheckSpelling** (_CustomDictionary_, _IgnoreUppercase_, _AlwaysSuggest_, _SpellLang_)

_expression_ A variable that represents a **[Range](excel.range(object).md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _CustomDictionary_|Optional| **Variant**|A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to have Microsoft Excel ignore words that are all uppercase. **False** to have Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.|
| _AlwaysSuggest_|Optional| **Variant**| **True** to have Excel display a list of suggested alternate spellings when an incorrect spelling is found. **False** to have Excel wait for you to input the correct spelling. If this argument is omitted, the current setting will be used.|
| _SpellLang_|Optional| **Variant**|The language of the dictionary being used. Can be one of the **[MsoLanguageID](Office.MsoLanguageID.md)** values.|

## Return value

Variant


## Remarks

This method has no return value; Microsoft Excel displays the **Spelling** dialog box.

To check headers, footers, and objects on a worksheet, use this method on a **[Worksheet](Excel.Worksheet.md)** object.

To check only cells and notes, use this method with the object returned by the **[Cells](Excel.Range.Cells.md)** property.




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]