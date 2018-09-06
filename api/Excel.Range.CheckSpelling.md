---
title: Range.CheckSpelling Method (Excel)
keywords: vbaxl10.chm144093
f1_keywords:
- vbaxl10.chm144093
ms.prod: excel
api_name:
- Excel.Range.CheckSpelling
ms.assetid: 22540515-0b0f-ce3c-0145-e46d477d1b5f
ms.date: 06/08/2017
---


# Range.CheckSpelling Method (Excel)

Checks the spelling of an object.


## Syntax

 _expression_. `CheckSpelling`( `_CustomDictionary_` , `_IgnoreUppercase_` , `_AlwaysSuggest_` , `_SpellLang_` )

 _expression_ A variable that represents a [Range](Excel.Range(Graph property).md) object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _CustomDictionary_|Optional| **Variant**|A string that indicates the file name of the custom dictionary to be examined if the word isn't found in the main dictionary. If this argument is omitted, the currently specified dictionary is used.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to have Microsoft Excel ignore words that are all uppercase. **False** to have Microsoft Excel check words that are all uppercase. If this argument is omitted, the current setting will be used.|
| _AlwaysSuggest_|Optional| **Variant**| **True** to have Microsoft Excel display a list of suggested alternate spellings when an incorrect spelling is found. **False** to have Microsoft Excel wait for you to input the correct spelling. If this argument is omitted, the current setting will be used.|
| _SpellLang_|Optional| **Variant**|The language of the dictionary being used. Can be one of the  **[MsoLanguageID](Office.MsoLanguageID.md)** values.|

### Return Value

Variant


## Remarks

This method has no return value; Microsoft Excel displays the  **Spelling** dialog box.

To check headers, footers, and objects on a worksheet, use this method on a  **[Worksheet](Excel.Worksheet.md)** object.

To check only cells and notes, use this method with the object returned by the  **[Cells](Excel.Range.Cells.md)** property.


## See also


[Range Object](Excel.Range(object).md)

