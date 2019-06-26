---
title: Window.WindowState property (Visio)
keywords: vis_sdr.chm11614670
f1_keywords:
- vis_sdr.chm11614670
ms.prod: visio
api_name:
- Visio.Window.WindowState
ms.assetid: 71578934-5d04-8e14-6d87-6871a31f9c4e
ms.date: 06/08/2017
localization_priority: Normal
---


# Window.WindowState property (Visio)

Gets or sets the state of a window. Read/write.


## Syntax

_expression_.**WindowState**

_expression_ A variable that represents a **[Window](Visio.Window.md)** object.


## Return value

Long


## Remarks

The **WindowState** property value can be a combination of the constants declared in the Visio type library in **[VisWindowStates](Visio.viswindowstates.md)**.


> [!NOTE] 
> The nFlags parameter to the **Add** method for the **Windows** collection can be composed of the various bits of **VisWindowStates**.

If you specify conflicting bits, only one bit is used. For example, if you specify both **visWSMaximized** and **visWSMinimized**, the window is maximized.

The **visWSVisible** flag is ignored when setting the state of a window with the **WindowState** property. It is used in calls to the **Add** method for the **Windows** collection. Use the **Visible** property of the window to show or hide it. The **visWSVisible** flag is available only when this property is read.

If your Visual Studio solution includes the [Microsoft.Office.Interop.Visio](https://docs.microsoft.com/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2019) reference, this property maps to the following types:


- **Microsoft.Office.Interop.Visio.IVWindow.WindowState**
    

## Example

This Microsoft Visual Basic for Applications (VBA) macro shows how to minimize the active drawing window.


```vb
Public Sub WindowState_Example() 
  
    Dim vsoWindow As Visio.Window      
 
    'Get the active window. 
    Set vsoWindow = ActiveWindow  
 
    'Minimize the active window. 
    vsoWindow.WindowState = visWSMinimized 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]