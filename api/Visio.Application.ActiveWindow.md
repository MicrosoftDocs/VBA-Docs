---
title: Application.ActiveWindow property (Visio)
keywords: vis_sdr.chm10013035
f1_keywords:
- vis_sdr.chm10013035
ms.prod: visio
api_name:
- Visio.Application.ActiveWindow
ms.assetid: 6da310fd-3fb1-618b-d80f-98ee1e45d5a2
ms.date: 06/24/2019
localization_priority: Normal
---


# Application.ActiveWindow property (Visio)

Returns the active **[Window](visio.window.md)** object. Read-only.


## Syntax

_expression_.**ActiveWindow**

_expression_ A variable that represents an **[Application](Visio.Application.md)** object.


## Return value

Window


## Remarks

The active window can be one of the following window types: Drawing, Stencil, ShapeSheet, Edit Icon, or a Drawing or Stencil window created by an add-on. 

The application's active window can only be an MDI frame window—it cannot be one of the floating, docked, or anchored windows. For a complete list of window types, see the **[Window.Type](Visio.Window.Type.md)** property.

If a window in an instance of Microsoft Visio is not active, the **ActiveWindow** property returns **Nothing**.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this property maps to the following types:

- **Microsoft.Office.Interop.Visio.IVApplication.ActiveWindow**
    

## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to get the active window without qualification from the Microsoft Office Visio global object, which is automatically available to VBA code that is part of the VBA project of a Visio document.

```vb
 
Public Sub ActiveWindow_Example() 
  
    Dim vsoWindow As Visio.Window  
 
    'Get the active window. 
    Set vsoWindow = ActiveWindow  
 
    'To verify that we got the active window, print its caption.  
    Debug.Print vsoWindow.Caption  
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]