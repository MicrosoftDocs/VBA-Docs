---
title: Windows object (Project)
keywords: vbapj.chm131357
f1_keywords:
- vbapj.chm131357
ms.prod: project-server
ms.assetid: 6fc70ece-0257-5565-907b-e0e7a6770980
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows object (Project)

Contains a collection of  **[Window](Project.Window.md)** objects. The **Windows** collection for the **Application** object contains all the windows in the application, whereas the **Windows** collection for the **Project** object contains only the windows in the specified project.
 


## Remarks


> [!NOTE] 
> The **Windows** collection is maintained for backward compatibility. We recommend that you use the **[Windows2](Project.windows2(object).md)** collection for all new development.
 


## Examples

 **Using the Window object**
 

 
Use  **Windows** ( _Index_), where  _Index_ is the window index number or window caption, to return a single **Window** object. The following example maximizes the first window in the window list.
 

 



```vb
Application.Windows(1).WindowState = pjMaximized
```

The window caption is the text shown in the title bar at the top of the window when the window is not maximized. The caption is also shown in the list of open files on the bottom of the **Windows** menu. Use the **[Caption](Project.Application.Caption.md)** property to set or return the window caption. Changing the window caption does not change the name of the project. The following example hides the window that contains the caption "Project1".
 

 



```vb
If Application.Windows(1).Caption = "Project1" Then  
    Application.Windows(1).Visible = False  
End If
```

 **Using the Windows collection**
 

 
Use the **[Windows](Project.Application.Windows.md)** property to return a **Windows** collection. The following example cascades all the windows that are currently displayed in Project .
 

 



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


## Properties



|Name|
|:-----|
|[ActiveWindow](Project.Windows.ActiveWindow.md)|
|[Application](Project.Windows.Application.md)|
|[Count](Project.Windows.Count.md)|
|[Item](Project.Windows.Item.md)|
|[Parent](Project.Windows.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]