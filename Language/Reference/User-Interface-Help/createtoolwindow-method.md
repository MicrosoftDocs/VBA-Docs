---
title: CreateToolWindow method
keywords: vbob6.chm100291
f1_keywords:
- vbob6.chm100291
ms.prod: office
api_name:
- Office.CreateToolWindow
ms.assetid: da49893c-8b04-5bda-f7ff-fd70a70a084f
ms.date: 12/06/2018
localization_priority: Normal
---


# CreateToolWindow method

Creates a new Tool window containing the indicated **UserDocument** object.

## Syntax

_object_.**CreateToolWindow** (_AddInInst, ProgID, Caption, GuidPosition, DocObj_) **As Window**

<br/>

The **CreateToolWindow** method syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _AddInInst_|Required. An add-in instance variable representing an add-in in the development environment.|
| _ProgID_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) representing the progID of the **UserDocument** object.|
| _Caption_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) containing the window caption.|
| _GuidPosition_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) containing a unique identifier for the window.|
| _DocObj_|Required. [Object](../../Glossary/vbe-glossary.md#object) representing a **UserDocument** object. This object will be set in the call to this function.|

## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]