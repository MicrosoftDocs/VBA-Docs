---
title: Windows2 object (Project)
keywords: vbapj.chm131358
f1_keywords:
- vbapj.chm131358
ms.prod: project-server
ms.assetid: a58383c6-12c7-81b3-10e8-81ba9180404c
ms.date: 06/08/2017
localization_priority: Normal
---


# Windows2 object (Project)

Represents a collection of windows in the application or project.
 


## Remarks


> [!NOTE] 
> We recommend that you use the **Windows2** collection instead of the **Windows** collection for all new development.
 

The **Windows2** collection for the **Application** object contains all the windows in the application, whereas the **Windows2** collection for the **Project** object contains only the windows in the specified project.
 

 
Development with the .NET Framework 4, or with external components or applications that automate Project, must use the **Windows2** object, not the **Windows** object. A Primary Interop Assembly (PIA) is required to interact correctly with the COM interface of Project if those components are to be signed. Microsoft Visual Studio generates an interop assembly based on the type library if no PIA is present, but the components then cannot be signed with a digital certificate. The PIA is installed with Project.
 

 

## Examples

 **Using the Windows2 collection**
 

 
Use the **[Windows2](Project.Application.Windows2.md)** property to return a **Windows2** collection.
 

 
The following example cascades all the windows that are currently displayed in Project.
 

 



```vb
With Application.Windows2  
    For I = 1 To .Count  
        .Item(I).Activate  
        .Item(I).Top = (I - 1) * 15  
        .Item(I).Left = (I - 1) * 15  
    Next I  
End With
```

Use the **[WindowNewWindow](Project.Application.WindowNewWindow.md)** method to create a new window and add it to the **Windows2** collection.
 

 
The following example creates a new window for the active project.
 

 



```vb
Application.WindowNewWindow
```

 **Using the Windows2 object**
 

 

## Using the Windows2 Object

Use **Windows2** (*Index* ), where*Index* is the window index number or window caption, to return a single **Window** object.
 

 
The following example maximizes the first window in the window list.
 

 



```vb
Application.Windows2(1).WindowState = pjMaximized
```

The window caption is the text shown in the title bar at the top of the window when the window is not maximized. The caption is also shown in the list of open files on the bottom of the **Windows** menu. Use the **[Caption](Project.Window.Caption.md)** property to set or return the window caption. Changing the window caption does not change the name of the project.
 

 
The following example hides the window that contains the caption "Project1".
 

 



```vb
If Application.Windows2(1).Caption = "Project1" Then  
    Application.Windows2(1).Visible = False  
End If
```


## Properties



|Name|
|:-----|
|[ActiveWindow](Project.Windows2.ActiveWindow.md)|
|[Application](Project.Windows2.Application.md)|
|[Count](Project.Windows2.Count.md)|
|[Item](Project.Windows2.Item.md)|
|[Parent](Project.Windows2.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]