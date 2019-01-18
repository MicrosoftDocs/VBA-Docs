---
title: CheckBox.EventProcPrefix property (Access)
keywords: vbaac10.chm10692
f1_keywords:
- vbaac10.chm10692
ms.prod: access
api_name:
- Access.CheckBox.EventProcPrefix
ms.assetid: 9ab63762-34fb-06f4-3b79-97471152c939
ms.date: 06/08/2017
localization_priority: Normal
---


# CheckBox.EventProcPrefix property (Access)

Gets or sets the prefix portion of an event procedure name. Read/write  **String**.


## Syntax

_expression_. `EventProcPrefix`

_expression_ A variable that represents a [CheckBox](Access.CheckBox.md) object.


## Remarks

For example, if you have a command button with an event procedure named Details_Click, the  **EventProcPrefix** property returns the string "Details".

Microsoft Access adds the prefix portion of an event procedure name to the event name with an underscore character (_).


## See also


[CheckBox Object](Access.CheckBox.md)

