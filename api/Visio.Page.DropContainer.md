---
title: Page.DropContainer method (Visio)
keywords: vis_sdr.chm10962135
f1_keywords:
- vis_sdr.chm10962135
ms.prod: visio
api_name:
- Visio.Page.DropContainer
ms.assetid: 14da134d-6a3f-25c3-37c4-eb8b51c213ab
ms.date: 06/08/2017
localization_priority: Normal
---


# Page.DropContainer method (Visio)

Creates a new container  **[Shape](Visio.Shape.md)** object on the page, places the container around the specified target shapes, and adds the target shapes to the container. Returns the container shape.


## Syntax

_expression_. `DropContainer`( `_ObjectToDrop_` , `_TargetShapes_` )

_expression_ A variable that represents a **[Page](Visio.Page.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectToDrop_|Required| **[UNKNOWN]**|The container shape to add to the page. Can be a  **[Master](Visio.Master.md)**, **[MasterShortcut](Visio.MasterShortcut.md)**, **Shape**, or **IDataObject** object.|
| _TargetShapes_|Required| **[UNKNOWN]**|The shapes that the container should contain. Can be a  **Shape** or a **[Selection](Visio.Selection.md)** object. The shapes or selection must already be on the page.|

## Return value

 **Shape**


## Remarks

To pass a  **Master** object for the _ObjectToDrop_ parameter, use the **[Documents.OpenEx](Visio.Documents.OpenEx.md)** method and the **[Application.GetBuiltInStencilFile](Visio.Application.GetBuiltInStencilFile.md)** method, passing it **visBuiltInStencilContainers**, to open the hidden, built-in container stencil. Then use the **[Masters.ItemU](Visio.Masters.ItemU.md)** property to get the particular container that you want from the stencil.

An  **IDataObject** that you pass for _ObjectToDrop_ must be provided by Microsoft Visio and must be in the same process space as Visio.

If  _ObjectToDrop_ is not a Visio object, or if it is not a container, Visio returns an Invalid Parameter error. If the value you pass is a shape that does not match the context of the method, Visio returns an Invalid Source error.

If the  _TargetShapes_ parameter is **Nothing**, Visio places the container shape at the center of the page, devoid of any target shapes. If the specified target shapes are not top-level members of the page, Visio returns an Invalid Parameter error.

The  **DropContainer** method corresponds to the **Insert Container** command in the Visio user interface. (On the **Insert** tab, click **Container**.)


## Example

The following Visual Basic for Applications (VBA) example shows how to use the  **DropContainer** method to add a container from the hidden, built-in container stencil to the active page to contain the selected shape or shapes. Before running this macro, be sure that there is a selected shape (or a selection of shapes) on the active page.


```vb
Public Sub DropContainer_Example()

    Dim vsoDocument As Visio.Document
    Set vsoDocument = Application.Documents.OpenEx(Application.GetBuiltInStencilFile(visBuiltInStencilContainers, visMSUS), visOpenHidden)
    Application.ActivePage.DropContainer vsoDocument.Masters.ItemU("Container 1"), Application.ActiveWindow.Selection
    vsoDocument.Close
    
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]