---
title: Collections (Visual Basic Add-In Model)
ms.prod: office
keywords: vbob6.chm1070949
f1_keywords:
- vbob6.chm1070949
- vbob6.chm1071203
- vbob6.chm1070951
- vbob6.chm1070956
- vbob6.chm1098895
- vbob6.chm1070948
- vbob6.chm102246
- vbob6.chm102053
- vbob6.chm1070945
ms.assetid: 45e5f192-c698-4805-9ba8-cbe52f313732
ms.date: 12/26/2018
localization_priority: Normal
---


# Collections (Visual Basic Add-In Model)

A collection is an object that contains a set of related objects. An object's position in the collection can change whenever a change occurs in the collection; therefore, the position of any specific object in the collection can vary.

The following sections describe the collections in the Visual Basic Add-In Model.

## AddIns

Returns a [collection](../../Glossary/vbe-glossary.md#collection) of [add-ins](../../Glossary/vbe-glossary.md#add-in) registered for VBA.

### Syntax

_object_.**AddIns**

### Remarks

The **AddIns** collection is accessed through the **[VBE](objects-visual-basic-add-in-model.md#vbe)** object. Every add-in listed in the Add-In Manager in VBE has an object in the **AddIns** collection. 

## CodePanes

Contains the active [code panes](../../Glossary/vbe-glossary.md#code-pane) in the **[VBE](objects-visual-basic-add-in-model.md#vbe)** object.

### Remarks

Use the **CodePanes** collection to access the open code panes in a [project](../../Glossary/vbe-glossary.md#project). 

You can use the **[Count](properties-visual-basic-add-in-model.md#count)** property to return the number of active code panes in a collection.

## CommandBars

Contains all of the [command bars](objects-visual-basic-add-in-model.md#commandbar) in a project, including command bars that support shortcut menus.

### Remarks

Use the **CommandBars** collection to enable add-ins to add command bars and [controls](../../Glossary/vbe-glossary.md#control), or to add controls to existing, built-in, command bars.

## LinkedWindows

Contains all linked [windows](objects-visual-basic-add-in-model.md#window) in a [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame).

### Remarks

Use the **LinkedWindows** collection to modify the [docked](../../Glossary/vbe-glossary.md#docked-window) and [linked](../../Glossary/vbe-glossary.md#linked-window) state of windows in the [development environment](../../Glossary/vbe-glossary.md#development-environment).

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

The **LinkedWindowFrame** property of the **[Window](objects-visual-basic-add-in-model.md#window)** object returns a **Window** object that has a valid **LinkedWindows** collection.

Linked window frames contain all windows that can be linked or docked. This includes all windows except code windows, [designers](../../Glossary/vbe-glossary.md#designer), the [Object Browser](../../Glossary/vbe-glossary.md#object-browser) window, and the Search and Replace window.

If all the panes from one linked window frame are moved to another window, the linked window frame with no panes is destroyed. However, if all the panes are removed from the main window, it isn't destroyed.

Use the **[Visible](properties-visual-basic-add-in-model.md#visible)** property to check or set the visibility of a window.

You can use the **[Add](../user-interface-help/add-method-vba-add-in-object-model.md)** method to add a window to the collection of currently linked windows. A window that is a pane in one linked window frame can be added to another linked window frame. Use the **[Remove](../user-interface-help/remove-method-vba-add-in-object-model.md)** method to remove a window from the collection of currently linked windows; this results in the window being unlinked or undocked.

The **LinkedWindows** collection is used to dock and undock windows from the main window frame.

## Properties

Represents the [properties](../../Glossary/vbe-glossary.md#property) of an object.

### Remarks

Use the **Properties** collection to access the properties displayed in the [Properties window](../user-interface-help/properties-window.md). For every property listed in the Properties window, there is a **[Property](objects-visual-basic-add-in-model.md#property)** object in the **Properties** collection.

## References

Represents the set of [references](objects-visual-basic-add-in-model.md#reference) in the project.

### Remarks

Use the **References** collection to add or remove references. The **References** collection is the same as the set of references selected in the **[References](../user-interface-help/references-dialog-box.md)** dialog box.

**See also** the **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** object.

## VBComponents

Represents the components contained in a project.

### Remarks

Use the **VBComponents** collection to access, add, or remove components in a project. A component can be a [form](../../Glossary/vbe-glossary.md#form), [module](../../Glossary/vbe-glossary.md#module), or [class](../../Glossary/vbe-glossary.md#class). The **VBComponents** collection is a standard collection that can be used in a **For...Each** block.

You can use the **[Parent](properties-visual-basic-add-in-model.md#parent)** property to return the project that the **VBComponents** collection is in.

For more information, see the **[VBComponents](properties-visual-basic-add-in-model.md#vbcomponents)** property and the **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** object.

## VBNewProjects

Represents all of the new projects in the development environment.

### Remarks

Use the **VBNewProjects** collection to access specific projects in an instance of the development environment. **VBNewProjects** is a standard collection that you can iterate through by using a **For...Each** block. 

## VBProjects

Represents all the projects that are open in the development environment.

### Remarks

Use the **VBProjects** collection to access specific projects in an instance of the development environment. **VBProjects** is a standard collection that can be used in a **For...Each** block.  

## Windows

Contains all open or permanent windows.

### Remarks

Use the **Windows** collection to access **[Window](objects-visual-basic-add-in-model.md#window)** objects.

The **Windows** collection has a fixed set of windows that are always available in the collection, such as the [Project window](../../Glossary/vbe-glossary.md#project-window), the Properties window, and a set of windows that represent all open code windows and designer windows. 

Opening a code or designer window adds a new member to the **Windows** collection. Closing a code or designer window removes a member from the **Windows** collection. Closing a permanent development environment window doesn't remove the corresponding object from this collection, but results in the window not being visible.

## See also

- [Drives collection](../user-interface-help/drives-collection.md)
- [Files collection](../user-interface-help/files-collection.md)
- [Folders collection](../user-interface-help/folders-collection.md)
- [Collections (Microsoft Forms)](../user-interface-help/objects-microsoft-forms.md)
- [Visual Basic Add-in Model reference](../user-interface-help/visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](../user-interface-help/visual-basic-language-reference.md)
- [Office client development reference](https://docs.microsoft.com/office/client-developer/office-client-development)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
