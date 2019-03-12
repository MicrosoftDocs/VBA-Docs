---
title: Properties (Visual Basic Add-In Model)
ms.prod: office
keywords: vbob6.chm100096
f1_keywords:
- vbob6.chm100096
ms.assetid: 278f8774-d259-4212-ba80-326841106aa1
ms.date: 12/26/2018 
localization_priority: Normal
---


# Properties (Visual Basic Add-In Model)

## ActiveCodePane

Returns the active or last active **[CodePane](objects-visual-basic-add-in-model.md#codepane)** object, or sets the active **CodePane** object. Read/write.

### Remarks

You can set the **ActiveCodePane** property to any valid **CodePane** object, as shown in the following example:

```vb
Set MyApp.VBE. ActiveCodePane = MyApp.VBE.CodePanes(1)

```

The preceding example sets the first [code pane](../../Glossary/vbe-glossary.md#code-pane) in a [collection](../../Glossary/vbe-glossary.md#collection) of code panes to be the active code pane. You can also activate a code pane by using the **[SetSelection](../user-interface-help/setselection-method-vba-add-in-object-model.md)** method.


## ActiveVBProject

Returns the active [project](../../Glossary/vbe-glossary.md#project) in the [Project window](../../Glossary/vbe-glossary.md#project-window). Read-only.

### Remarks

The **ActiveVBProject** property returns the project that is selected in the Project window or the project in which the components are selected. In the latter case, the project itself isn't necessarily selected. Whether or not the project is explicitly selected, there is always an active project.

## ActiveWindow

Returns the active window in the [development environment](../../Glossary/vbe-glossary.md#development-environment). Read-only.

### Remarks

When more than one window is open in the development environment, the **ActiveWindow** property setting is the window with the [focus](../../Glossary/vbe-glossary.md#focus). If the main window has the focus, **ActiveWindow** returns **[Nothing](../user-interface-help/nothing-keyword.md)**.

## AddIns

Returns a collection which add-ins can use to register their automation components into the extensibility object model.

### Syntax

_object_.**AddIns**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## BuildFileName

Sets or returns the DLL name that will be used when the project is built.

### Syntax

_object_.**BuildFileName**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## BuiltIn

Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value indicating whether the **[Reference](objects-visual-basic-add-in-model.md#reference)** object is a default reference that can't be removed. Read-only.

The **BuiltIn** property returns these values:

|Value|Description|
|:-----|:-----|
|**True**|The reference is a default reference that can't be removed.|
|**False**|The reference isn't a default reference; it can be removed.|

## Caption

Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the title of the active **[Window](objects-visual-basic-add-in-model.md#window)**. Read-only.

### Remarks

The title of the active window is the text displayed in the window's title bar.

## CodeModule

Returns an object representing the code behind the component. Read-only.

### Remarks

The **CodeModule** property returns **[Nothing](../user-interface-help/nothing-keyword.md)** if the component doesn't have a [code module](../../Glossary/vbe-glossary.md#code-module) associated with it.

> [!NOTE] 
> The **[CodePane](objects-visual-basic-add-in-model.md#codepane)** object represents a visible code window. A given component can have several **CodePane** objects. 
> 
> The **[CodeModule](objects-visual-basic-add-in-model.md#codemodule)** object represents the code within a component. A component can only have one **CodeModule** object.

## CodePane

Returns a **[CodePane](objects-visual-basic-add-in-model.md#codepane)** object. Read-only.

### Remarks

If a [code pane](../../Glossary/vbe-glossary.md#code-pane) exists, it becomes the active code pane, and the window that contains it becomes the active window. If a code pane doesn't exist for the [module](../../Glossary/vbe-glossary.md#module), the **CodePane** property creates one.

## CodePanes

Returns the [collection](../../Glossary/vbe-glossary.md#collection) of active **[CodePane](objects-visual-basic-add-in-model.md#codepane)** objects. Read-only.

## CodePaneView

Returns a value indicating whether the **[CodePane](objects-visual-basic-add-in-model.md#codepane)** is in Procedure view or Full Module view. Read-only.

The **CodePaneView** property returns these values:

|Constant|Description|
|:-----|:-----|
|**vbext_cv_ProcedureView**|The specified code pane is in Procedure view.|
|**vbext_cv_FullModuleView**|The specified [project](../../Glossary/vbe-glossary.md#project) is in Full Module view.|

## Collection

Returns the [collection](../../Glossary/vbe-glossary.md#collection) that contains the object you are working with. Read-only.

### Remarks

Most objects in this object model have either a **[Parent](#parent)** property or a **Collection** property that points to the object's parent object.

Use the **Collection** property to access the [properties](../../Glossary/vbe-glossary.md#property), [methods](../../Glossary/vbe-glossary.md#method), and [controls](../../Glossary/vbe-glossary.md#control) of the collection to which the object belongs.

## CommandBarEvents

Returns the **[CommandBarEvents](objects-visual-basic-add-in-model.md#commandbarevents)** object. Read-only.

### Settings

The setting for the [argument](../../Glossary/vbe-glossary.md#argument) you pass to the **CommandBarEvents** property is:

|Argument|Description|
|:-----|:-----|
| _vbcontrol_|Must be an object of type **[CommandBarControl](../../../api/office.commandbarcontrol.md)**.|

### Remarks

Use the **CommandBarEvents** property to return an [event source object](../../Glossary/vbe-glossary.md#event-source-object) that triggers an event when a command bar button is clicked. 

The argument passed to the **CommandBarEvents** property is the command bar control for which the **[Click](events-visual-basic-add-in-model.md#click-event)** event will be triggered.

## CommandBars

Contains all of the command bars in a project, including command bars that support shortcut menus.

**See also** [Menus and commands](../menus-commands.md) and [Toolbars](../toolbars.md).

## Connect

Returns or sets the connected state of an add-in.

### Remarks

Returns **True** if the add-in is registered and currently connected (active).

Returns **False** if the add-in is registered, but not connected (inactive).

## Count

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the number of items in a [collection](../../Glossary/vbe-glossary.md#collection). Read-only.

## CountOfDeclarationLines

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the number of lines of code in the Declarations section of a [code module](objects-visual-basic-add-in-model.md#codemodule). Read-only.

## CountOfLines

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the number of lines of code in a [code module](objects-visual-basic-add-in-model.md#codemodule). Read-only.

## CountOfVisibleLines

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the number of lines visible in a [code pane](objects-visual-basic-add-in-model.md#codepane). Read-only.

## Description

Returns or sets a [string expression](../../Glossary/vbe-glossary.md#string-expression) containing a descriptive string associated with an object. For the **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** object, read/write; for the **[Reference](objects-visual-basic-add-in-model.md#reference)** object, read-only.

### Remarks

For the **VBProject** object, the **Description** property returns or sets a descriptive string associated with the active [project](../../Glossary/vbe-glossary.md#project).

For the **Reference** object, the **Description** property returns the descriptive name of the reference.


## Designer

Returns the object that enables you to access the design characteristics of a component.

### Remarks

If the object has an open [designer](../../Glossary/vbe-glossary.md#designer), the **Designer** property returns the open designer; otherwise, a new designer is created. 

The designer is a characteristic of certain **[VBComponent](objects-visual-basic-add-in-model.md#vbcomponent)** objects. For example, when you create certain types of **VBComponent** objects, a designer is created along with the object. A component can have only one designer, and it's always the same designer. 

The **Designer** property enables you to access a component-specific object. In some cases, such as in [standard modules](../../Glossary/vbe-glossary.md#standard-module) and [class modules](../../Glossary/vbe-glossary.md#class-module), a designer isn't created because that type of **VBComponent** object doesn't support a designer.

The **Designer** property returns **[Nothing](../user-interface-help/nothing-keyword.md)** if the **VBComponent** object doesn't have a designer.


## DesignerID

Read-only property that returns the [ProgID](#progid) of an ActiveX designer.


## Events

Supplies properties that enable add-ins to connect to all [events](objects-visual-basic-add-in-model.md#events) in Visual Basic for Applications.

### Syntax

_object_.**Events**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## FileName

Returns the full path name of the project file or host document.

### Syntax

_object_.**Filename**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

### Remarks

Projects have no name other than the file name. The path name returned is always provided as an absolute path (for example, "c:\projects\myproject.vba"), even if it is shown as a relative path (such as "..\projects\myproject.vba").

## FullPath

Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the path and file name of the referenced [type library](../../Glossary/vbe-glossary.md#type-library). Read-only.

## GUID

Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the class identifier of an object. Read-only.

## HasOpenDesigner

Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value indicating whether the **[VBComponent](objects-visual-basic-add-in-model.md#vbcomponent)** object has an open [designer](../../Glossary/vbe-glossary.md#designer). Read-only.

The **HasOpenDesigner** property returns these values:

|Value|Description|
|:-----|:-----|
|**True**|The **VBComponent** object has an open Design window.|
|**False**|The **VBComponent** object doesn't have an open Design window.|

## Height

Returns or sets a [Single](../../Glossary/vbe-glossary.md#single-data-type) containing the height of the window in [twips](../../Glossary/vbe-glossary.md#twip). Read/write.

### Remarks

Changing the **Height** property setting of a [linked window](../../Glossary/vbe-glossary.md#linked-window) or [docked window](../../Glossary/vbe-glossary.md#docked-window) has no effect as long as the window remains linked or docked.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## HelpContextID

Returns or sets a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the context ID for a topic in a Microsoft Windows Help file. Read/write.

## HelpFile

Returns or sets a [String](../../Glossary/vbe-glossary.md#string-data-type) specifying the Microsoft Windows Help file for a [project](../../Glossary/vbe-glossary.md#project). Read/write.

## IndexedValue

Returns or sets a value for a member of a [property](../../Glossary/vbe-glossary.md#property) that is an indexed list or an [array](../../Glossary/vbe-glossary.md#array).

### Remarks

The value returned or set by the **IndexedValue** property is an [expression](../../Glossary/vbe-glossary.md#expression) that evaluates to a type that is accepted by the object. For a property that is an indexed list or array, you must use the **IndexedValue** property instead of the **[Value](#value)** property. An indexed list is a [numeric expression](../../Glossary/vbe-glossary.md#numeric-expression) specifying index position. Values in indexed lists are set or returned with a single index.

**IndexedValue** accepts up to 4 indices. The number of indices accepted by **IndexedValue** is the value returned by the **[NumIndices](#numindices)** property. The **IndexedValue** property is used only if the value of the **NumIndices** property is greater than zero. 


## IsBroken

Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value indicating whether the **[Reference](objects-visual-basic-add-in-model.md#reference)** object points to a valid reference in the [registry](../../Glossary/vbe-glossary.md#registry). Read-only.

The **IsBroken** property returns these values:

|Value|Description|
|:-----|:-----|
|**True**|The **Reference** object no longer points to a valid reference in the registry.|
|**False**|The **Reference** object points to a valid reference in the registry.|

## Left

Returns or sets a [Single](../../Glossary/vbe-glossary.md#single-data-type) containing the location of the left edge of the window on the screen in [twips](../../Glossary/vbe-glossary.md#twip). Read/write.

### Remarks

The value returned by the **Left** property depends on whether the window is [linked](../../Glossary/vbe-glossary.md#linked-window) or [docked](../../Glossary/vbe-glossary.md#docked-window).

> [!NOTE] 
> Changing the **Left** property setting of a linked or docked window has no effect as long as the window remains linked or docked.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## Lines

Returns a string containing the specified number of lines of code.

### Syntax

_object_.**Lines** (_startline_, _count_) **As String**

The **Lines** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _startline_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the line number in which to start.|
| _count_|Required. A **Long** specifying the number of lines you want to return.|

### Remarks

The [line numbers](../../Glossary/vbe-glossary.md#line-number) in a [code module](../../Glossary/vbe-glossary.md#code-module) begin at 1.

## LinkedWindowFrame

Returns the **[Window](objects-visual-basic-add-in-model.md#window)** object representing the frame that contains the window. Read-only.

### Remarks

The **LinkedWindowFrame** property enables you to access the object representing the [linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame), which has properties distinct from the window or windows it contains. If the window isn't linked, the **LinkedWindowFrame** property returns **[Nothing](../user-interface-help/nothing-keyword.md)**.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## LinkedWindows

Returns the collection of all linked [windows](objects-visual-basic-add-in-model.md#window) contained in a linked window frame. Read-only.

### Remarks

The **LinkedWindows** property is an accessor property (that is, a property that returns an object of the same type as the property name).

## MainWindow

Returns a **[Window](objects-visual-basic-add-in-model.md#window)** object representing the main window of the Visual Basic [development environment](../../Glossary/vbe-glossary.md#development-environment). Read-only.

### Remarks

You can use the **Window** object returned by the **MainWindow** property to add or remove [docked windows](../../Glossary/vbe-glossary.md#docked-window), and to maximize, minimize, hide, or restore the main window of the Visual Basic development environment.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## Major

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) containing the major version number of the referenced [type library](../../Glossary/vbe-glossary.md#type-library). Read-only.

### Remarks

The number returned by the **Major** property corresponds to the major version number stored in the type library to which you have set the reference.

## Minor

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) indicating the minor version number of the referenced [type library](../../Glossary/vbe-glossary.md#type-library). Read-only.

### Remarks

The number returned by the **Minor** property corresponds to the minor version number stored in the type library to which you have set the reference.

## Mode

Returns a value containing the mode of the specified [project](../../Glossary/vbe-glossary.md#project). Read-only.

The **Mode** property returns these values:

|Constant|Description|
|:-----|:-----|
|**vbext_vm_Run**|The specified project is in run mode.|
|**vbext_vm_Break**|The specified project is in break mode.|
|**vbext_vm_Design**|The specified project is in design mode.|

## Name

Returns or sets a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the name used in code to identify an object. For the **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** object and the **[VBComponent](objects-visual-basic-add-in-model.md#vbcomponent)** object, read/write. For the **[Property](objects-visual-basic-add-in-model.md#property)** object and the **[Reference](objects-visual-basic-add-in-model.md#reference)** object, read-only.

### Remarks

The following table describes how the **Name** property setting applies to different objects.

|Object|Result of using Name property setting|
|:-----|:-----|
|**VBProject**|Returns or sets the name of the active [project](../../Glossary/vbe-glossary.md#project).|
|**VBComponent**|Returns or sets the name of the component. An error occurs if you try to set the **Name** property to a name already being used or an invalid name.|
|**Property**|Returns the name of the property as it appears in the **Property Browser**. This is the value used to index the **[Properties](collections-visual-basic-add-in-model.md#properties)** collection. The name can't be set.|
|**Reference**|Returns the name of the reference in code. The name can't be set.|

The default name for new objects is the type of object plus a unique integer. For example, the first new **Form** object is Form1, a new **Form** object is Form1, and the third **TextBox** control that you create on a form is TextBox3.

An object's **Name** property must start with a letter and can be a maximum of 40 characters. It can include numbers and underline (_) characters, but can't include punctuation or spaces. 

[Forms](../../Glossary/vbe-glossary.md#form) and [modules](../../Glossary/vbe-glossary.md#module) can't have the same name as another public object such as **Clipboard**, **Screen**, or **App**. Although the **Name** property setting can be a [keyword](../../Glossary/vbe-glossary.md#keyword), property name, or the name of another object, this can create conflicts in your code.

## NumIndices

Returns the number of indices on the [property](../../Glossary/vbe-glossary.md#property) returned by the **[Property](objects-visual-basic-add-in-model.md#property)** object.

### Remarks

The value of the **NumIndices** property can be an integer from 0&ndash;4. For most properties, **NumIndices** returns 0. Conventionally indexed properties return 1. Property [arrays](../../Glossary/vbe-glossary.md#array) might return 2.

## Object

Returns or sets the value of an object returned by a [property](../../Glossary/vbe-glossary.md#property). Read/write.

### Remarks

If a **[Property](objects-visual-basic-add-in-model.md#property)** object returns an object, you must use the **Object** property to return or set the value of that object.

## Parent

Returns the object or [collection](../../Glossary/vbe-glossary.md#collection) that contains another object or collection. Read-only.

### Remarks

Most objects have either a **Parent** property or a **[Collection](#collection)** property that points to the object's parent object in this object model. The **Collection** property is used if the parent object is a collection.

Use the **Parent** property to access the [properties](../../Glossary/vbe-glossary.md#property), [methods](../../Glossary/vbe-glossary.md#method), and [controls](../../Glossary/vbe-glossary.md#control) of an object's parent object.

**See also** [CodeModule object](objects-visual-basic-add-in-model.md#codemodule)

## ProcBodyLine

Returns the first line of a [procedure](../../Glossary/vbe-glossary.md#procedure).

### Syntax

_object_.**ProcBodyLine** (_procname_, _prockind_) **As Long**

<br/>

The **ProcBodyLine** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _procname_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the name of the procedure.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](../../Glossary/vbe-glossary.md#property-procedure) can have multiple representations in the [module](../../Glossary/vbe-glossary.md#module), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is, **[Sub](../user-interface-help/sub-statement.md)** and **[Function](../user-interface-help/function-statement.md)** procedures) use **vbext_pk_Proc**.|

<br/>

You can use one of the following [constants](../../Glossary/vbe-glossary.md#constant) for the _prockind_ [argument](../../Glossary/vbe-glossary.md#argument).

|Constant|Description|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|

### Remarks

The first line of a procedure is the line on which the **Sub**, **Function**, or **[Property](../user-interface-help/property-get-statement.md)** statement appears.

## ProcCountLines

Returns the number of lines in the specified [procedure](../../Glossary/vbe-glossary.md#procedure).

### Syntax

_object_.**ProcCountLines** (_procname_, _prockind_) **As Long**

<br/>

The **ProcCountLines** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _procname_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the name of the procedure.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](../../Glossary/vbe-glossary.md#property-procedure) can have multiple representations in the [module](../../Glossary/vbe-glossary.md#module), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is, **[Sub](../user-interface-help/sub-statement.md)** and **[Function](../user-interface-help/function-statement.md)** procedures) use **vbext_pk_Proc**.|

<br/>

You can use one of the following [constants](../../Glossary/vbe-glossary.md#constant) for the _prockind_ [argument](../../Glossary/vbe-glossary.md#argument).

|Constant|Description|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|

### Remarks

The **ProcCountLines** property returns the count of all blank or comment lines preceding the procedure declaration and, if the procedure is the last procedure in a [code module](../../Glossary/vbe-glossary.md#code-module), any blank lines following the procedure.


## ProcOfLine

Returns the name of the [procedure](../../Glossary/vbe-glossary.md#procedure) that the specified line is in.

### Syntax

_object_.**ProcOfLine** (_line_, _prockind_) **As String**

<br/>

The **ProcOfLine** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _line_|Required. A [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the line to check.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](../../Glossary/vbe-glossary.md#property-procedure) can have multiple representations in the [module](../../Glossary/vbe-glossary.md#module), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is, **[Sub](../user-interface-help/sub-statement.md)** and **[Function](../user-interface-help/function-statement.md)** procedures) use **vbext_pk_Proc**.|

<br/>

You can use one of the following [constants](../../Glossary/vbe-glossary.md#constant) for the _prockind_ [argument](../../Glossary/vbe-glossary.md#argument).

|Constant|Description|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a procedure that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|

### Remarks

A line is within a procedure if it's a blank line or comment line preceding the procedure declaration and, if the procedure is the last procedure in a [code module](objects-visual-basic-add-in-model.md#codemodule), a blank line or lines following the procedure.

## ProcStartLine

Returns the line at which the specified [procedure](../../Glossary/vbe-glossary.md#procedure) begins.

### Syntax

_object_.**ProcStartLine** (_procname_, _prockind_) **As Long**

<br/>

The **ProcStartLine** syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. An [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.|
| _procname_|Required. A [String](../../Glossary/vbe-glossary.md#string-data-type) containing the name of the procedure.|
| _prockind_|Required. Specifies the kind of procedure to locate. Because [property procedures](../../Glossary/vbe-glossary.md#property-procedure) can have multiple representations in the [module](../../Glossary/vbe-glossary.md#module), you must specify the kind of procedure you want to locate. All procedures other than property procedures (that is, **[Sub](../user-interface-help/sub-statement.md)** and **[Function](../user-interface-help/function-statement.md)** procedures) use **vbext_pk_Proc**.|

<br/>

You can use one of the following [constants](../../Glossary/vbe-glossary.md#constant) for the _prockind_ [argument](../../Glossary/vbe-glossary.md#argument).

|Constant|Description|
|:-----|:-----|
|**vbext_pk_Get**|Specifies a [procedure](../../Glossary/vbe-glossary.md#procedure) that returns the value of a property.|
|**vbext_pk_Let**|Specifies a procedure that assigns a value to a property.|
|**vbext_pk_Set**|Specifies a procedure that sets a reference to an object.|
|**vbext_pk_Proc**|Specifies all procedures other than property procedures.|

### Remarks

A procedure starts at the first line below the **[End Sub](../user-interface-help/end-statement.md)** statement of the preceding procedure. If the procedure is the first procedure, it starts at the end of the general Declarations section.

## ProgID

Returns the ProgID (programmatic ID) for the control represented by the **VBControl** object.

### Syntax

_object_.**ProgID**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Properties

Returns the properties of an object. Read-only.

### Remarks

The **Properties** property is an accessor property (that is, a property that returns an object of the same type as the property name).

## Protection

Returns a value indicating the state of protection of a [project](../../Glossary/vbe-glossary.md#project). Read-only.

The **Protection** property returns these values:

|Constant|Description|
|:-----|:-----|
|**vbext_pp_locked**|The specified project is locked.|
|**vbext_pp_none**|The specified project isn't protected.|

## References

Returns the set of [references](objects-visual-basic-add-in-model.md#reference) in a project. Read-only.

### Remarks

The **References** property is an accessor property (that is, a property that returns an object of the same type as the property name).

## ReferencesEvents

Returns the **[ReferencesEvents](objects-visual-basic-add-in-model.md#referencesevents)** object. Read-only.

### Settings

The setting for the [argument](../../Glossary/vbe-glossary.md#argument) you pass to the **ReferencesEvents** property is:

|Argument|Description|
|:-----|:-----|
| _vbproject_|If _vbproject_ points to **[Nothing](../user-interface-help/nothing-keyword.md)**, the object that is returned will supply events for the **[References](collections-visual-basic-add-in-model.md#references)** collections of all **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** objects in the **[VBProjects](collections-visual-basic-add-in-model.md#vbprojects)** collection.<br/><br/>If _vbproject_ points to a valid **VBProject** object, the object that is returned will supply events for only the **References** collection for that [project](../../Glossary/vbe-glossary.md#project).|

### Remarks

The **ReferencesEvents** property takes an argument and returns an [event source object](../../Glossary/vbe-glossary.md#event-source-object). The **ReferencesEvents** object is the source for events that are triggered when references are added or removed.

## Saved

Returns a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value indicating whether the object was edited since the last time it was saved. Read/write.

The **Saved** property returns these values:

|Value|Description|
|:-----|:-----|
|**True**|The object has not been edited since the last time it was saved.|
|**False**|The object has been edited since the last time it was saved.|

### Remarks

The **[SaveAs](../user-interface-help/saveas-method-vba-add-in-object-model.md)** method sets the **Saved** property to **True**.

> [!NOTE] 
> If you set the **Saved** property to **False** in code, it returns **False**, and the object is marked as if it were edited since the last time it was saved.

## SelectedVBComponent

Returns the selected component. Read-only.

### Remarks

The **SelectedVBComponent** property returns the selected component in the [Project window](../../Glossary/vbe-glossary.md#project-window). If the selected item in the Project window isn't a component, **SelectedVBComponent** returns **[Nothing](../user-interface-help/nothing-keyword.md)**.

**See also** the **[VBE](objects-visual-basic-add-in-model.md#vbe)** object.

## Top

Returns or sets a [Single](../../Glossary/vbe-glossary.md#single-data-type) specifying the location of the top of the window on the screen in [twips](../../Glossary/vbe-glossary.md#twip). Read/write.

### Remarks

The value returned by the **Top** property depends on whether the window is [docked](../../Glossary/vbe-glossary.md#docked-window), [linked](../../Glossary/vbe-glossary.md#linked-window), or in docking view.

> [!NOTE] 
> Changing the **Top** property setting of a linked or docked window has no effect as long as the window remains linked or docked.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## TopLine

Returns a [Long](../../Glossary/vbe-glossary.md#long-data-type) specifying the line number of the line at the top of the [code pane](objects-visual-basic-add-in-model.md#codepane), or sets the line showing at the top of the code pane. Read/write.

### Remarks

Use the **TopLine** property to return or set the line showing at the top of the code pane. For example, if you want line 25 to be the first line showing in a code pane, set the **TopLine** property to 25.

The **TopLine** property setting must be a positive number. If the **TopLine** property setting is greater than the actual number of lines in the code pane, the setting will be the last line in the code pane.

## Type

Returns a numeric or string value containing the type of object. Read-only.

The **Type** property settings for the **[Window](objects-visual-basic-add-in-model.md#window)** object are described in the following table.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbext_wt_CodeWindow**|0|[Code window](../user-interface-help/code-window.md)|
|**vbext_wt_Designer**|1|[Designer](../../Glossary/vbe-glossary.md#designer)|
|**vbext_wt_Browser**|2|**[Object Browser](../user-interface-help/object-browser.md)**|
|**vbext_wt_Immediate**|5|[Immediate window](../user-interface-help/immediate-window.md)|
|**vbext_wt_ProjectWindow**|6|[Project window](../user-interface-help/project-explorer.md)|
|**vbext_wt_PropertyWindow**|7|[Properties window](../user-interface-help/properties-window.md)|
|**vbext_wt_Find**|8|**[Find](../user-interface-help/find-dialog-box.md)** dialog box|
|**vbext_wt_FindReplace**|9|**[Search and Replace](../user-interface-help/replace-dialog-box.md)** dialog box|
|**vbext_wt_LinkedWindowFrame**|11|[Linked window frame](../../Glossary/vbe-glossary.md#linked-window-frame)|
|**vbext_wt_MainWindow**|12|Main window|
|**vbext_wt_Watch**|3|[Watch window](../user-interface-help/watch-window.md)|
|**vbext_wt_Locals**|4|[Locals window](../user-interface-help/locals-window.md)|
|**vbext_wt_Toolbox**|10|**[Toolbox](../user-interface-help/toolbox.md)**|
|**vbext_wt_ToolWindow**|15|Tool window|

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

<br/>

The **Type** property settings for the **[VBComponent](objects-visual-basic-add-in-model.md#vbcomponent)** object are described in the following table.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbext_ct_StdModule**|1|[Standard module](../../Glossary/vbe-glossary.md#standard-module)|
|**vbext_ct_ClassModule**|2|[Class module](../../Glossary/vbe-glossary.md#class-module)|
|**vbext_ct_MSForm**|3|Microsoft Form|
|**vbext_ct_ActiveXDesigner**|11|ActiveX Designer|
|**vbext_ct_Document**|100|Document Module|

<br/>

The **Type** property settings for the **[Reference](objects-visual-basic-add-in-model.md#reference)** object are described in the following table.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbext_rk_TypeLib**|0|[Type library](../../Glossary/vbe-glossary.md#type-library)|
|**vbext_rk_Project**|1|[Project](../../Glossary/vbe-glossary.md#project)|

<br/>

The **Type** property settings for the **[VBProject](objects-visual-basic-add-in-model.md#vbproject)** object are described in the following table.

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbext_pt_HostProject**|100|Host Project|
|**vbext_pt_StandAlone**|101|Standalone Project|

## Value

Returns or sets a [Variant](../../Glossary/vbe-glossary.md#variant-data-type) specifying the value of the [property](../../Glossary/vbe-glossary.md#property). Read/write.

### Remarks

Because the **Value** property returns a **Variant**, you can access any property. To access a list, use the **[IndexedValue](#indexedvalue)** property.

If the property that the **[Property](objects-visual-basic-add-in-model.md#property)** object represents is read/write, the **Value** property is read/write. If the property is read-only, attempting to set the **Value** property causes an error. If the property is write-only, attempting to return the **Value** property causes an error. 

The **Value** property is the default property for the **Property** object.

## VBComponents

Returns a collection of the components contained in a project.

### Remarks

Use the **[VBComponents](collections-visual-basic-add-in-model.md#vbcomponents)** collection to access, add, or remove components in a project. A component can be a [form](../../Glossary/vbe-glossary.md#form), [module](../../Glossary/vbe-glossary.md#module), or [class](../../Glossary/vbe-glossary.md#class). The **VBComponents** collection is a standard [collection](../../Glossary/vbe-glossary.md#collection) that can be used in a **For… Each** block.

You can use the **[Parent](#parent)** property to return the project that the **VBComponents** collection is in.

In Visual Basic for Applications, you can use the **[Import](../user-interface-help/import-method-vba-add-in-object-model.md)** method to add a component to a project from a file.

For more information, see **[VBComponent](objects-visual-basic-add-in-model.md#vbcomponent)** object and **[SelectedVBComponent](#selectedvbcomponent)**  property.

## VBE

Returns the root of the **[VBE](objects-visual-basic-add-in-model.md#vbe)** object. Read-only.

### Remarks

All objects have a **VBE** property that points to the root of the **VBE** object.

## VBProjects

Returns the **[VBProjects](collections-visual-basic-add-in-model.md#vbprojects)** collection, which represents all of the projects currently open in the Visual Basic IDE.

### Syntax

_object_.**VBProjects**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## Version

Returns a [String](../../Glossary/vbe-glossary.md#string-data-type) containing the version of Visual Basic for Applications that the application is using. Read-only.

### Remarks

The **Version** property value is a string beginning with one or two digits, a period, and two digits; the rest of the string is undefined and may contain text or numbers.

## Visible

For the **[Window](objects-visual-basic-add-in-model.md#window)** object, returns or sets a [Boolean](../../Glossary/vbe-glossary.md#boolean-data-type) value that specifies the visibility of a window. Read/write. 

For the **[CodePane](objects-visual-basic-add-in-model.md#codepane)** object, returns a **Boolean** value that indicates whether the [code pane](../../Glossary/vbe-glossary.md#code-pane) is visible in the window. Read-only.

The **Visible** property returns these values:

|Value|Description|
|:-----|:-----|
|**True**|(Default) Object is visible.|
|**False**|Object is hidden.|

## Width

Returns or sets a [Single](../../Glossary/vbe-glossary.md#single-data-type) containing the width of the window in [twips](../../Glossary/vbe-glossary.md#twip). Read/write.

### Remarks

Changing the **Width** property setting of a [linked window](../../Glossary/vbe-glossary.md#linked-window) or [docked window](../../Glossary/vbe-glossary.md#docked-window) has no effect as long as the window remains linked or docked.

> [!IMPORTANT] 
> Objects, properties, and methods for controlling linked windows, linked window frames, and docked windows are included on the Macintosh for compatibility with code written in Windows. However, these language elements will generate run-time errors when run on the Macintosh.

## Window

Returns the window in which the [code pane](objects-visual-basic-add-in-model.md#codepane) is displayed. Read-only.

## Windows

Returns the **[Window](objects-visual-basic-add-in-model.md#window)** object, which represents a window in the Visual Basic IDE.

### Syntax

_object_.**Window**

The _object_ placeholder represents an [object expression](../../Glossary/vbe-glossary.md#object-expression) that evaluates to an object in the **Applies To** list.

## WindowState

Returns or sets a numeric value specifying the visual state of the **[Window](objects-visual-basic-add-in-model.md#window)**. Read/write.

### Settings

The **WindowState** property returns or sets the following values:

|Constant|Value|Description|
|:-----|:-----|:-----|
|**vbext_ws_Normal**|0|(Default) Normal|
|**vbext_ws_Minimize**|1|Minimized (minimized to an icon)|
|**vbext_ws_Maximize**|2|Maximized (enlarged to maximum size)|


    
## See also

- [Properties (Microsoft Forms)](../user-interface-help/properties-microsoft-forms.md)
- [Objects, methods, and properties (Visual Basic for Applications)](../objects-visual-basic-for-applications.md)
- [Visual Basic Add-in Model reference](../user-interface-help/visual-basic-add-in-model-reference.md)
- [Visual Basic language reference](../user-interface-help/visual-basic-language-reference.md)
- [Office client development reference](https://docs.microsoft.com/office/client-developer/office-client-development)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
