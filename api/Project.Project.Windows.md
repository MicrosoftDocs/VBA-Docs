---
title: Project.Windows property (Project)
ms.prod: project-server
api_name:
- Project.Project.Windows
ms.assetid: 6096d9b0-c6ba-3e48-9265-f3e8448d293f
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Windows property (Project)

Gets a **[Windows](Project.windows(object).md)** collection representing the open windows in the project. Read-only **Windows**.


## Syntax

_expression_.**Windows**

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **Windows** property duplicates the **Windows2** property, except that it returns a **Windows** collection rather than a **Windows2** collection. The **[Windows2](Project.Application.Windows2.md)** property and **[Windows2](Project.windows2(object).md)** collection object are recommended for all new development both in VBA and for external applications developed with the .NET Framework. The **Windows** property and **Windows** collection are maintained for backward compatibility with existing applications.


## Example

The following example cascades all the open windows in the active project.


```vb
Sub CascadeWindows() 
 Dim I As Integer 
 
 ActiveWindow.WindowState = pjNormal ' Restore the window. 
 
 With ActiveProject.Windows 
 For I = 1 To .Count 
 .Item(I).Activate 
 .Item(I).Top = (I - 1) * 15 
 .Item(I).Left = (I - 1) * 15 
 Next I 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]