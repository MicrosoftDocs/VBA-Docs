---
title: CreateToolWindow Method
keywords: vbob6.chm100291
f1_keywords:
- vbob6.chm100291
ms.prod: office
api_name:
- Office.CreateToolWindow
ms.assetid: da49893c-8b04-5bda-f7ff-fd70a70a084f
ms.date: 06/08/2017
---


# CreateToolWindow Method



Creates a new Tool window containing the indicated  **UserDocument** object.

## Syntax

_object_. **CreateToolWindow (**_AddInInst, ProgID, Caption, GuidPosition, DocObj_**) As Window**
The  **CreateToolWindow** method syntax has these parts:


|Part|Description|
|:-----|:-----|
| _object_|An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the Applies To list.|
| _AddInInst_|Required. An add-in instance variable representing an add-in in the development environment.|
| _ProgID_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) representing the progID of the **UserDocument** object.|
| _Caption_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) containing the window caption.|
| _GuidPosition_|Required. [String](../../Glossary/vbe-glossary.md#string-data-type) containing a unique identifier for the window.|
| _DocObj_|Required. [Object](../../Glossary/vbe-glossary.md#object) representing a **UserDocument** object. This object will be set in the call to this function.|

