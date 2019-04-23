---
title: Events (Visual Basic Add-In Model)
ms.prod: office
keywords: vbob6.chm1098932
f1_keywords:
- vbob6.chm1098932
- vbob6.chm1098927
- vbob6.chm100150
ms.assetid: ae90ce4d-7f61-4e7d-a4ab-7cf78028281a
ms.date: 12/26/2018
localization_priority: Normal
---


# Events (Visual Basic Add-In Model)

## Click event

Occurs when the **OnAction** [property](../../Glossary/vbe-glossary.md#property) of a corresponding command bar control is set.

### Syntax

**Sub**_object_**\_Click** (**ByVal** _ctrl_ **As Object**, **ByRef** _handled_ **As Boolean**, **ByRef** _canceldefault_ **As Boolean**)

<br/>

The **Click** event syntax has these [named arguments](../../Glossary/vbe-glossary.md#named-argument):

|Part|Description|
|:-----|:-----|
|_ctrl_|Required; [Object](../../Glossary/vbe-glossary.md#object). Specifies the object that is the source of the **Click** event.|
|_handled_|Required; [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type). If **True**, other [add-ins](../../Glossary/vbe-glossary.md#add-in) should handle the event. If **False**, the action of the command bar item has not been handled.|
|_canceldefault_|Required; **Boolean**. If **True**, default behavior is performed unless canceled by a downstream add-in. If **False**, default behavior is not performed unless restored by a downstream add-in.|

### Remarks

The **Click** event is specific to the **[CommandBarEvents](objects-visual-basic-add-in-model.md#commandbarevents)** object. Use a [variable](../../Glossary/vbe-glossary.md#variable) declared by using the **WithEvents** keyword to receive the **Click** event for a **CommandBar** control. This variable should be set to the return value of the **[CommandBarEvents](properties-visual-basic-add-in-model.md#commandbarevents)** property of the **[Events](objects-visual-basic-add-in-model.md#events)** object. 

The **CommandBarEvents** property takes the **CommandBar** control as an [argument](../../Glossary/vbe-glossary.md#argument). When the **CommandBar** control is clicked (for the variable you declared by using the **WithEvents** keyword), the code is executed.

## ItemAdded event

Occurs after a reference is added.

### Syntax

**Sub**_object_**\_ItemAdded** (**ByVal** _item_ **As Reference**)

The required _item_ [argument](../../Glossary/vbe-glossary.md#argument) specifies the item that was added.

### Remarks

The **ItemAdded** event occurs when a **[Reference](objects-visual-basic-add-in-model.md#reference)** is added to the **[References](collections-visual-basic-add-in-model.md#references)** collection.

## ItemRemoved event

Occurs after a **Reference** is removed from a [project](../../Glossary/vbe-glossary.md#project).

### Syntax

**Sub**_object_**\_ItemRemoved** (**ByVal** _item_ **As Reference**)

The required _item_ [argument](../../Glossary/vbe-glossary.md#argument) specifies the **Reference** that was removed.

## See also

- [ReferencesEvents object](objects-visual-basic-add-in-model.md#referencesevents)
- [Events (Microsoft Forms)](../user-interface-help/events-microsoft-forms.md)
- [Events (Visual Basic for Applications)](../events-visual-basic-for-applications.md)
- [Visual Basic Add-in Model reference](../user-interface-help/visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](../user-interface-help/visual-basic-language-reference.md)
- [Office client development reference](https://docs.microsoft.com/office/client-developer/office-client-development)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]