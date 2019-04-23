---
title: Objects (Visual Basic Add-In Model)
ms.prod: office
keywords: vbob6.chm100000
f1_keywords:
- vbob6.chm100000
- vbob6.chm1071201
- vbob6.chm102045
- vbob6.chm104053
- vbob6.chm1071195
- vbob6.chm1070944
- vbob6.chm1071196
- vbob6.chm1092845
- vbob6.chm104030
- vbob6.chm1093161
- vbob6.chm100108
- vbob6.chm102256
ms.assetid: d92a32be-3e40-4ce2-ba11-fa797840984a
ms.date: 12/26/2018
localization_priority: Normal
---


# Objects (Visual Basic Add-In Model)

## AddIn

The **AddIn** object provides information about an add-in to other add-ins.

### Syntax

_object_.**AddIn**

### Remarks

An **AddIn** object is created for every add-in that appears in the [Add-In Manager](../user-interface-help/add-in-manager-dialog-box.md).

## CodeModule

Represents the code behind a component, such as a [form](../../Glossary/vbe-glossary.md#form), [class](../../Glossary/vbe-glossary.md#class), or [document](../../Glossary/vbe-glossary.md#document).

### Remarks

You use the **CodeModule** object to modify (add, delete, or edit) the code associated with a component. Each component is associated with one **CodeModule** object. However, a **CodeModule** object can be associated with multiple [code panes](#codepane).

The methods associated with the **CodeModule** object enable you to manipulate and return information about the code text on a line-by-line basis. For example, you can use the **[AddFromString](../user-interface-help/addfromstring-method-vba-add-in-object-model.md)** method to add text to the [module](../../Glossary/vbe-glossary.md#module). **AddFromString** places the text just above the first [procedure](../../Glossary/vbe-glossary.md#procedure) in the module or places the text at the end of the module if there are no procedures.

Use the **[Parent](properties-visual-basic-add-in-model.md#parent)** property to return the **[VBComponent](#vbcomponent)** object associated with a [code module](../../Glossary/vbe-glossary.md#code-module).

## CodePane

Represents a [code pane](../../Glossary/vbe-glossary.md#code-pane).

### Remarks

Use the **CodePane** object to manipulate the position of visible text or the text selection displayed in the code pane.

You can use the **[Show](../user-interface-help/show-method-vba-add-in-object-model.md)** method to make the code pane you specify visible. 

Use the **[SetSelection](../user-interface-help/setselection-method-vba-add-in-object-model.md)** method to set the selection in a code pane. 

Use the **[GetSelection](../user-interface-help/getselection-method-vba-add-in-object-model.md)** method to return the location of the selection in a code pane.

## CommandBar

The **CommandBar** object contains other **CommandBar** objects, which can act as either buttons or menu commands.

### Syntax

**CommandBar**

## CommandBarEvents

Returned by the **[CommandBarEvents](properties-visual-basic-add-in-model.md#commandbarevents)** property. The **CommandBarEvents** object triggers an event when a [control](../../Glossary/vbe-glossary.md#control) on the command bar is clicked.

### Remarks

The **CommandBarEvents** object is returned by the **CommandBarEvents** property of the **[Events](#events)** object. 

The object that is returned has one event in its interface, the **[Click](events-visual-basic-add-in-model.md#click-event)** event. You can handle this event by using the **WithEvents** object declaration.

## Events

Supplies [properties](../../Glossary/vbe-glossary.md#property) that enable [add-ins](../../Glossary/vbe-glossary.md#add-in) to connect to all events in Visual Basic for Applications.

### Remarks

The **Events** object provides properties that return [event source objects](../../Glossary/vbe-glossary.md#event-source-object). Use the properties to return event source objects that notify you of changes in the Visual Basic for Applications environment.

The properties of the **Events** object return objects of the same type as the property name. For example, the **CommandBarEvents** property returns the **CommandBarEvents** object.

## Property

Represents the [properties](../../Glossary/vbe-glossary.md#property) of an object that are visible in the [Properties window](../../Glossary/vbe-glossary.md#properties-window) for any given component.

### Remarks

Use the **[Value](properties-visual-basic-add-in-model.md#value)** property of the **Property** object to return or set the value of a property of a component.

At a minimum, all components have a **[Name](properties-visual-basic-add-in-model.md#name)** property. The **Value** property returns a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) of the appropriate type. If the value returned is an object, the **Value** property returns the **[Properties](collections-visual-basic-add-in-model.md#properties)** collection that contains **Property** objects representing the individual properties of the object. You can access each of the **Property** objects by using the **[Item](../user-interface-help/item-method-vba-add-in-object-model.md)** method on the returned **Properties** collection.

If the value returned by the **Property** object is an object, you can use the **[Object](properties-visual-basic-add-in-model.md#object)** property to set the **Property** object to a new object.

## Reference

Represents a reference to a [type library](../../Glossary/vbe-glossary.md#type-library) or a [project](../../Glossary/vbe-glossary.md#project).

### Remarks

Use the **Reference** object to verify whether a reference is still valid.

The **[IsBroken](properties-visual-basic-add-in-model.md#isbroken)** property returns **True** if the reference no longer points to a valid reference. 

The **[BuiltIn](properties-visual-basic-add-in-model.md#builtin)** property returns **True** if the reference is a default reference that can't be moved or removed. 

Use the **[Name](properties-visual-basic-add-in-model.md#name)** property to determine if the reference you want to add or remove is the correct one.

**See also** the **[Description](properties-visual-basic-add-in-model.md#description)** and **[Type](properties-visual-basic-add-in-model.md#type)** properties.

## ReferencesEvents

Returned by the **[ReferencesEvents](properties-visual-basic-add-in-model.md#referencesevents)** property.

### Remarks

The **ReferencesEvents** object is the source of events that occur when a reference is added to or removed from a [project](../../Glossary/vbe-glossary.md#project). 

The **[ItemAdded](events-visual-basic-add-in-model.md#itemadded-event)** event is triggered after a reference is added to a project.

The **[ItemRemoved](events-visual-basic-add-in-model.md#itemremoved-event)** event is triggered after a reference is removed from a project.

## VBComponent

Represents a component, such as a [class module](../../Glossary/vbe-glossary.md#class-module) or [standard module](../../Glossary/vbe-glossary.md#standard-module), contained in a [project](../../Glossary/vbe-glossary.md#project).

### Remarks

Use the **VBComponent** object to access the **[CodeModule](#codemodule)** object associated with a component or to change a component's property settings.

You can use the **[Type](properties-visual-basic-add-in-model.md#type)** property to find out what type of component the **VBComponent** object refers to. 

Use the **[Collection](properties-visual-basic-add-in-model.md#collection)** property to find out what [collection](../../Glossary/vbe-glossary.md#collection) the component is in.

## VBE

The root object that contains all other [objects](../../Glossary/vbe-glossary.md#object) and [collections](../../Glossary/vbe-glossary.md#collection) represented in Visual Basic for Applications.

### Remarks

You can use the following [collections](collections-visual-basic-add-in-model.md) to access the objects contained in the **VBE** object:

- Use the **VBProjects** collection to access the collection of [projects](../../Glossary/vbe-glossary.md#project).
    
- Use the **AddIns** collection to access the collection of add-ins.
    
- Use the **Windows** collection to access the collection of windows.
    
- Use the **CodePanes** collection to access the collection of [code panes](../../Glossary/vbe-glossary.md#code-pane).
    
- Use the **CommandBars** collection to access the collection of command bars.
    

Use the **[Events](#events)** object to access properties that enable [add-ins](../../Glossary/vbe-glossary.md#add-in) to connect to all events in Visual Basic for Applications. The properties of the **Events** object return objects of the same type as the property name. For example, the **CommandBarEvents** property returns the **CommandBarEvents** object.

You can use the **[SelectedVBComponent](properties-visual-basic-add-in-model.md#selectedvbcomponent)** property to return the active component. The active component is the component that is being tracked in the [Project window](../../Glossary/vbe-glossary.md#project-window). If the selected item in the Project window isn't a component, **SelectedVBComponent** returns **[Nothing](../user-interface-help/nothing-keyword.md)**.

> [!NOTE] 
> All objects in this object model have a **[VBE](properties-visual-basic-add-in-model.md#vbe)** property that points to the **VBE** object.

## VBProject

Represents a [project](../../Glossary/vbe-glossary.md#project).

### Remarks

Use the **VBProject** object to set [properties](../../Glossary/vbe-glossary.md#property) for the project, and to access the **[VBComponents](collections-visual-basic-add-in-model.md#vbcomponents)** and **[References](collections-visual-basic-add-in-model.md#references)** collections.

## Window

Represents a window in the [development environment](../../Glossary/vbe-glossary.md#development-environment).

### Remarks

Use the **Window** object to show, hide, or position windows.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

You can use the **[Close](../user-interface-help/close-method-vba-add-in-object-model.md)** method to close a window in the **[Windows](collections-visual-basic-add-in-model.md#windows)** collection. The **Close** method affects different types of windows as follows:

|Window|Result of using Close method|
|:-----|:-----|
|Code window|Removes the window from the **Windows** collection.|
|[Designer](../../Glossary/vbe-glossary.md#designer)|Removes the window from the **Windows** collection.|
|**Window** objects of type [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame)|Windows become unlinked separate windows.|

> [!NOTE] 
> Using the **Close** method with code windows and designers actually closes the window. Setting the **[Visible](properties-visual-basic-add-in-model.md#visible)** property to **False** hides the window but doesn't close the window. Using the **Close** method with development environment windows, such as the [Project window](../../Glossary/vbe-glossary.md#project-window) or [Properties window](../../Glossary/vbe-glossary.md#properties-window), is the same as setting the **Visible** property to **False**.

You can use the **[SetFocus](../user-interface-help/setfocus-method-vba-add-in-object-model.md)** method to move the [focus](../../Glossary/vbe-glossary.md#focus) to a window.

You can use the **[Visible](properties-visual-basic-add-in-model.md#visible)** property to return or set the visibility of a window.

To find out what type of window you are working with, you can use the **[Type](properties-visual-basic-add-in-model.md#type)** property. If you have more than one window of a type, for example, multiple designers, you can use the **[Caption](properties-visual-basic-add-in-model.md#caption)** property to determine the window you are working with. 

You can also find the window you want to work with by using the **DesignerWindow** property of the **[VBComponent](#vbcomponent)** object or the **Window** property of the **[CodePane](#codepane)** object.

## See also

- [Objects (Microsoft Forms)](../user-interface-help/objects-microsoft-forms.md)
- [Objects and collections (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic Add-in Model reference](../user-interface-help/visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](../user-interface-help/visual-basic-language-reference.md)
- [Office client development reference](https://docs.microsoft.com/office/client-developer/office-client-development)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]