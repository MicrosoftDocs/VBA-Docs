---
title: Add method (VBA Add-In Object Model)
keywords: vbob6.chm1014017
f1_keywords:
- vbob6.chm1014017
ms.prod: office
ms.assetid: 95f4b970-0b0a-a41d-6a7b-8ede6626da67
ms.date: 12/06/2018
localization_priority: Normal
---


# Add method (VBA Add-In Object Model)

Adds an object to a [collection](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md).

## Syntax

_object_.**Add** (_component_) 

<br/>

The **Add** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _component_|Required. For the **LinkedWindows** collection, an object. For the **VBComponents** collection, an enumerated [constant](../../Glossary/vbe-glossary.md#constant) representing a [class module](../../Glossary/vbe-glossary.md#class-module), a form, or a [standard module](../../Glossary/vbe-glossary.md#standard-module). For the **VBProjects** collection, an enumerated constant representing a project type.|

<br/>

You can use one of the following constants for the _component_ argument:

|Constant|Description|
|:-----|:-----|
|**vbext_ct_ClassModule**|Adds a class module to the collection.|
|**vbext_ct_MSForm**|Adds a form to the collection.|
|**vbext_ct_StdModule**|Adds a standard module to the collection.|
|**vbext_pt_StandAlone**|Adds a standalone project to the collection.|

## Remarks

For the **LinkedWindows** collection, the **Add** method adds a window to the collection of currently [linked windows](../../Glossary/vbe-glossary.md#linked-window).

> [!NOTE] 
> You can add a window that is a pane in one [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame) to another linked window frame; the window is simply moved from one pane to the other. If the linked window frame that the window was moved from no longer contains any panes, it's destroyed.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements generate run-time errors when run on the Macintosh.

For the **VBComponents** collection, the **Add** method creates a new standard component, adds it to the [project](../../Glossary/vbe-glossary.md#project), and returns a **[VBComponent](vbcomponent-object-vba-add-in-object-model.md)** object. 

For the **LinkedWindows** collection, the **Add** method returns **[Nothing](nothing-keyword.md)**.

For the **VBProjects** collection, the **Add** method returns a **[VBProject](vbproject-object-vba-add-in-object-model.md)** object and adds a project to the **VBProjects** collection.



## See also

- [Collections (Visual Basic Add-In Model)](../visual-basic-add-in-model/collections-visual-basic-add-in-model.md)
- [Visual Basic Add-in Model reference](visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](visual-basic-language-reference.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]