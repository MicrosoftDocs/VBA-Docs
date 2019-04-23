---
title: Find method (VBA Add-In Object Model)
keywords: vbob6.chm1098972
f1_keywords:
- vbob6.chm1098972
ms.prod: office
ms.assetid: cf7a4b4e-89e7-91ea-2f9b-880384cd3339
ms.date: 12/06/2018
localization_priority: Normal
---


# Find method (VBA Add-In Object Model)

Searches the active [module](../../Glossary/vbe-glossary.md#module) for a specified string.

## Syntax

_object_.**Find** (_target_, _startline_, _startcol_, _endline_, _endcol_, [ _wholeword_ ], [ _matchcase_ ], [ _patternsearch_ ]) **As Boolean**

<br/>

The **Find** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _target_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the text or pattern you want to find.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the line at which you want to start the search; will be set to the line of the match if one is found. The first line is number 1.|
| _startcol_|Required. A **Long** specifying the column at which you want to start the search; will be set to the column containing the match if one is found. The first column is 1.|
| _endline_|Required. A **Long** specifying the last line of the match if one is found. The last line may be specified as -1.|
| _endcol_|Required. A **Long** specifying the last line of the match if one is found. The last column may be designated as -1.|
| _wholeword_|Optional. A [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value specifying whether to only match whole words. If **True**, only matches whole words. **False** is the default.|
| _matchcase_|Optional. A **Boolean** value specifying whether to match case. If **True**, the search is case sensitive. **False** is the default.|
| _patternsearch_|Optional. A **Boolean** value specifying whether or not the target string is a regular expression pattern. If **True**, the target string is a regular expression pattern. **False** is the default.|

## Remarks

**Find** returns **True** if a match is found and **False** if a match isn't found.

The _matchcase_ and _patternsearch_ [arguments](../../Glossary/vbe-glossary.md#argument) are mutually exclusive; if both arguments are passed as **True**, an error occurs.

The content of the **[Find](find-dialog-box.md)** dialog box isn't affected by the **Find** method.

The specified range of lines and columns is inclusive, so a search can find the pattern on the specified last line if _endcol_ is supplied as either -1 or the length of the line.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]