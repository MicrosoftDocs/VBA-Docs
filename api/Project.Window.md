---
title: Window object (Project)
keywords: vbapj.chm131356
f1_keywords:
- vbapj.chm131356
ms.prod: project-server
api_name:
- Project.Window
ms.assetid: b5dcb82d-1f5a-1334-0f03-3e23d3b9d940
ms.date: 06/08/2017
localization_priority: Normal
---


# Window object (Project)

Represents a window in the application or project. The **Window** object is a member of the **[Windows](Project.windows(object).md)** collection.
 


## Remarks


> [!NOTE] 
> The **Windows** collection is maintained for backward compatibility. We recommended that you use the **[Windows2](Project.windows2(object).md)** collection for all new development.
 

The **Application.Windows** collection contains all the windows in the application, whereas the **Project.Windows** collection contains only the windows in the specified project.
 

 

## Examples

 **Using the Window object**
 

 
Use  **Windows** (*Index* ), where*Index* is the window index number or window caption, to return a single **Window** object. The following example maximizes the first window in the window list.
 

 



```vb
Application.Windows(1).WindowState = pjMaximized
```

The window caption is the text shown in the title bar at the top of the window when the window is not maximized. The caption is also shown in the list of open files on the bottom of the **Windows** menu. Use the **[Caption](Project.Window.Caption.md)** property to set or return the window caption. Changing the window caption does not change the name of the project. The following example hides the window that contains the caption "Project1".
 

 



```vb
If Application.Windows(1).Caption = "Project1" Then
    Application.Windows(1).Visible = False
End If
```

 **Using the Windows collection**
 

 
Use the **[Windows](Project.Application.Windows.md)** property to return a **Windows** collection. The following example cascades all the windows that are currently displayed in Project.
 

 



```vb
With Application.Windows
    For I = 1 To .Count
        .Item(I).Activate
        .Item(I).Top = (I - 1) * 15
        .Item(I).Left = (I - 1) * 15
    Next I
End With
```

Use the **[WindowNewWindow](Project.Application.WindowNewWindow.md)** method to create a new window and add it to the collection. The following example creates a new window for the active project.
 

 



```vb
Application.WindowNewWindow
```


## Methods



|Name|
|:-----|
|[Activate](Project.Window.Activate.md)|
|[Close](Project.Window.Close.md)|
|[Refresh](Project.Window.Refresh.md)|
|[WebBrowserControlFrame](Project.Window.WebBrowserControlFrame.md)|
|[WebBrowserControlWindow](Project.Window.WebBrowserControlWindow.md)|

## Properties



|Name|
|:-----|
|[ActivePane](Project.Window.ActivePane.md)|
|[Application](Project.Window.Application.md)|
|[BottomPane](Project.Window.BottomPane.md)|
|[Caption](Project.Window.Caption.md)|
|[Height](Project.Window.Height.md)|
|[Index](Project.Window.Index.md)|
|[Left](Project.Window.Left.md)|
|[Parent](Project.Window.Parent.md)|
|[Top](Project.Window.Top.md)|
|[TopPane](Project.Window.TopPane.md)|
|[Visible](Project.Window.Visible.md)|
|[Width](Project.Window.Width.md)|
|[WindowState](Project.Window.WindowState.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]