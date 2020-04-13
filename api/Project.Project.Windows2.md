---
title: Project.Windows2 property (Project)
ms.prod: project-server
api_name:
- Project.Project.Windows2
ms.assetid: 0f10c401-d09b-82db-60ed-0f2b03b82656
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Windows2 property (Project)

Gets a **[Windows2](Project.windows2(object).md)** collection representing the open windows in the project. Read-only **Windows2**.


## Syntax

_expression_. `Windows2`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **Windows2** property is recommended, in place of the **Windows** property, for all new development in VBA and external applications developed with the .NET Framework.


## Example

The following example cascades all the open windows in the active project.


```vb
Sub CascadeWindows() 
 Dim I As Integer 
 
 ActiveWindow.WindowState = pjNormal ' Restore the window. 
 
 With ActiveProject.Windows2 
 For I = 1 To .Count 
 .Item(I).Activate 
 .Item(I).Top = (I - 1) * 15 
 .Item(I).Left = (I - 1) * 15 
 Next I 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]