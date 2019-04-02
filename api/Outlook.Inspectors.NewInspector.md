---
title: Inspectors.NewInspector event (Outlook)
keywords: vbaol11.chm312
f1_keywords:
- vbaol11.chm312
ms.prod: outlook
api_name:
- Outlook.Inspectors.NewInspector
ms.assetid: 945fb1a6-262f-da0d-16c6-bc27193505ac
ms.date: 06/08/2017
localization_priority: Normal
---


# Inspectors.NewInspector event (Outlook)

Occurs whenever a new inspector window is opened, either as a result of user action or through program code. 


## Syntax

_expression_. `NewInspector`( `_Inspector_` )

_expression_ A variable that represents an [Inspectors](Outlook.Inspectors.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Inspector_|Required| **[Inspector](Outlook.Inspector.md)**|The inspector that was opened.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).

The event occurs after the new  **Inspector** object is created but before the inspector window appears.


## See also


[Inspectors Object](Outlook.Inspectors.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]