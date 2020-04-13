---
title: Application.GetSpellingSuggestions method (Word)
keywords: vbawd10.chm158335303
f1_keywords:
- vbawd10.chm158335303
ms.prod: word
api_name:
- Word.Application.GetSpellingSuggestions
ms.assetid: 9ddf8aa8-10cc-8dd3-bc87-cdd5ccd214b5
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GetSpellingSuggestions method (Word)

Returns a  **[SpellingSuggestions](Word.spellingsuggestions.md)** collection that represents the words suggested as spelling replacements for a given word.


## Syntax

_expression_. `GetSpellingSuggestions`( `_Word_` , `_CustomDictionary_` , `_IgnoreUppercase_` , `_MainDictionary_` , `_SuggestionMode_` , `_CustomDictionary2_` , `_CustomDictionary3_` , `_CustomDictionary4_` , `_CustomDictionary5_` , `_CustomDictionary6_` , `_CustomDictionary7_` , `_CustomDictionary8_` , `_CustomDictionary9_` , `_CustomDictionary10_` )

_expression_ Required. A variable that represents an **[Application](Word.Application.md)** object. 


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Word_|Required| **String**|The word whose spelling is to be checked.|
| _IgnoreUppercase_|Optional| **Variant**| **True** to ignore words in all uppercase letters. If this argument is omitted, the current value of the **[IgnoreUppercase](Word.Options.IgnoreUppercase.md)** property is used.|
| _SuggestionMode_|Optional| **Variant**|Specifies the way Word makes spelling suggestions. Can be one of the following  **[WdSpellingWordType](Word.WdSpellingWordType.md)** constants: **wdAnagram**, **wdSpellword**, or **wdWildcard**. The default value is **WdSpellword**.|

## Remarks

If the word is spelled correctly, the **Count** property of the **[SpellingSuggestions](Word.spellingsuggestions.md)** object returns 0 (zero).


## See also


[Application Object](Word.Application.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]