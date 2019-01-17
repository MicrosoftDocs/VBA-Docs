---
title: SetSelection method (VBA Add-In Object Model)
keywords: vbob6.chm104035
f1_keywords:
- vbob6.chm104035
ms.prod: office
ms.assetid: c6408c78-b41e-e0d7-1817-41f887ce2d50
ms.date: 12/06/2018
localization_priority: Normal
---


# SetSelection method (VBA Add-In Object Model)

Sets the selection in the [code pane](../../Glossary/vbe-glossary.md#code-pane).

## Syntax

_object_.**SetSelection** (_startline_, _startcol_, _endline_, _endcol_)

<br/>

The **SetSelection** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the first line of the selection.|
| _startcol_|Required. A **Long** specifying the first column of the selection.|
| _endline_|Required. A **Long** specifying the last line of the selection.|
| _endcol_|Required. A **Long** specifying the last column of the selection.|

## See also

- [CodePane object](../visual-basic-add-in-model/objects-visual-basic-add-in-model.md#codepane)
- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]