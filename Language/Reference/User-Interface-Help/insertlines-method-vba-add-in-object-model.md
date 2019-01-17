---
title: InsertLines method (VBA Add-In Object Model)
keywords: vbob6.chm1098975
f1_keywords:
- vbob6.chm1098975
ms.prod: office
ms.assetid: 6a719fb8-cb52-6a18-c0dc-a8cd09a4814d
ms.date: 12/06/2018
localization_priority: Normal
---


# InsertLines method (VBA Add-In Object Model)

Inserts a line or lines of code at a specified location in a block of code.

## Syntax

_object_.**InsertLines** (_line_, _code_)

<br/>

The **InsertLines** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _line_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the location at which you want to insert the code.|
| _code_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the code you want to insert.|

## Remarks

If the text you insert by using the **InsertLines** method is carriage return-linefeed delimited, it will be inserted as consecutive lines.

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]