---
title: Form.BeforeScreenTip event (Access)
keywords: vbaac10.chm13678
f1_keywords:
- vbaac10.chm13678
ms.prod: access
api_name:
- Access.Form.BeforeScreenTip
ms.assetid: 08e67747-9023-e880-c246-1aa9e9c447ed
ms.date: 03/08/2019
localization_priority: Normal
---


# Form.BeforeScreenTip event (Access)

Occurs before a ScreenTip is displayed for an element in a PivotChart view or PivotTable view.


## Syntax

_expression_.**BeforeScreenTip** (_ScreenTipText_, _SourceObject_)

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ScreenTipText_|Required|**Object**|Set the **Value** property of this object to the ScreenTip that you want to display. Changing this argument to an empty string effectively hides the ScreenTip.|
| _SourceObject_|Required|**Object**|The object that generates the ScreenTip.|

## Return value

Nothing




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]