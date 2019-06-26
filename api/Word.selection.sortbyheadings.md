---
title: Selection.SortByHeadings method (Word)
keywords: vbawd10.chm158663698
f1_keywords:
- vbawd10.chm158663698
ms.prod: word
ms.assetid: fc38c337-b658-7b8d-2191-2ee98a93b48e
ms.date: 06/08/2017
localization_priority: Normal
---


# Selection.SortByHeadings method (Word)

Sorts the headings in the specified selection.


## Syntax

_expression_.**SortByHeadings** (_SortFieldType_, _SortOrder_, _CaseSensitive_, _BidiSort_, _IgnoreThe_, _IgnoreKashida_, _IgnoreDiacritics_, _IgnoreHe_, _LanguageID_)

_expression_ A variable that represents a **[Selection](Word.Selection.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _SortFieldType_|Optional|**Variant**|The sort field type to use. Can be one of the **[WdSortFieldType](Word.WdSortFieldType.md)** constants. The default value is **wdSortFieldAlphanumeric**. Depending on the language support (U.S. English, for example) that you have selected or installed, some of these constants may not be available to you.|
| _SortOrder_|Optional|**Variant**|The sorting order to use. Can be one of the **[WdSortOrder](Word.WdSortOrder.md)** constants.|
| _CaseSensitive_|Optional|**Variant**| **True** to sort with case sensitivity. The default value is **False**.|
| _BidiSort_|Optional|**Variant**| **True** to sort based on right-to-left language rules. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreThe_|Optional|**Variant**| **True** to ignore the Arabic character alef lam when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreKashida_|Optional|**Variant**| **True** to ignore kashidas when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreDiacritics_|Optional|**Variant**| **True** to ignore bidirectional control characters when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _IgnoreHe_|Optional|**Variant**| **True** to ignore the Hebrew character he when sorting right-to-left language text. Depending on the language support (U.S. English, for example) that you have selected or installed, this parameter may not be available to you.|
| _LanguageID_|Optional|**Variant**|Specifies the sorting language. Can be one of the **[WdLanguageID](Word.WdLanguageID.md)** constants.|

## Return value

**VOID**




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]