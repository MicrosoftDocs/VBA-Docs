---
title: Click event (VBA Add-In Object Model)
keywords: vbob6.chm1098932
f1_keywords:
- vbob6.chm1098932
ms.prod: office
ms.assetid: ac72bf41-e582-be84-d204-96545e8db71e
ms.date: 12/06/2018
---


# Click event (VBA Add-In Object Model)

Occurs when the **OnAction** [property](../../Glossary/vbe-glossary.md#property) of a corresponding command bar control is set.

## Syntax

**Sub**_object_**\_Click** (**ByVal** _ctrl_ **As Object**, **ByRef** _handled_ **As Boolean**, **ByRef** _canceldefault_ **As Boolean**)

<br/>

The **Click** event syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_ctrl_|Required; [Object](../../Glossary/vbe-glossary.md#object). Specifies the object that is the source of the **Click** event.|
|_handled_|Required; [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type). If **True**, other [add-ins](../../Glossary/vbe-glossary.md#add-in) should handle the event. If **False**, the action of the command bar item has not been handled.|
|_canceldefault_|Required; **Boolean**. If **True**, default behavior is performed unless canceled by a downstream add-in. If **False**, default behavior is not performed unless restored by a downstream add-in.|

## Remarks

The **Click** event is specific to the **[CommandBarEvents](commandbarevents-object-vba-add-in-object-model.md)** object. Use a [variable](../../Glossary/vbe-glossary.md#variable) declared by using the **WithEvents** keyword to receive the **Click** event for a **CommandBar** control. This variable should be set to the return value of the **[CommandBarEvents](commandbarevents-property-vba-add-in-object-model.md)** property of the **[Events](events-object-vba-add-in-object-model.md)** object. 

The **CommandBarEvents** property takes the **CommandBar** control as an [argument](../../Glossary/vbe-glossary.md#argument). When the **CommandBar** control is clicked (for the variable you declared by using the **WithEvents** keyword), the code is executed.

## See also

- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)