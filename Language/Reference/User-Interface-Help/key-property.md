---
title: Key property (Visual Basic for Applications)
keywords: vblr6.chm2181947
f1_keywords:
- vblr6.chm2181947
ms.prod: office
api_name:
- Office.Key
ms.assetid: 6b2d19f0-9729-7c36-fc22-bde7d6366fc8
ms.date: 12/19/2018
localization_priority: Normal
---


# Key property

Sets a _key_ in a **[Dictionary](dictionary-object.md)** object.

## Syntax

_object_.**Key** (_key_) = _newkey_

<br/>

The **Key** property has the following parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. Always the name of a **Dictionary** object.|
| _key_|Required. The _key_ value being changed.|
| _newkey_|Required. New value that replaces the specified _key_.|

## Remarks

If _key_ is not found when changing a _key_, a [run-time error](../../Glossary/vbe-glossary.md#run-time-error) will occur.

## See also

- [Objects (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]