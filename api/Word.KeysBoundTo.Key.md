---
title: KeysBoundTo.Key method (Word)
keywords: vbawd10.chm160890881
f1_keywords:
- vbawd10.chm160890881
ms.prod: word
api_name:
- Word.KeysBoundTo.Key
ms.assetid: efaef450-7d8d-0099-2420-07ae44c6bfa1
ms.date: 06/08/2017
localization_priority: Normal
---


# KeysBoundTo.Key method (Word)

Returns a  **KeyBinding** object that represents the specified custom key combination.


## Syntax

_expression_.**Key** (_KeyCode_, _KeyCode2_)

_expression_ A variable that represents a '[KeysBoundTo](Word.keysboundto.md)' object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _KeyCode_|Required| **Long**|A key you specify by using one of the **WdKey** constants.|
| _KeyCode2_|Optional| **Variant**|A second key you specify by using one of the **WdKey** constants.|

## Return value

KeyBinding


## Remarks

If the key combination doesn't exist, this method returns  **Nothing**.

You can use the **BuildKeyCode** method to create the KeyCode or KeyCode2 argument.


## See also


[KeysBoundTo Collection Object](Word.keysboundto.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]