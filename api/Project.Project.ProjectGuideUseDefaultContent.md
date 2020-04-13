---
title: Project.ProjectGuideUseDefaultContent property (Project)
keywords: vbapj.chm131090
f1_keywords:
- vbapj.chm131090
ms.prod: project-server
api_name:
- Project.Project.ProjectGuideUseDefaultContent
ms.assetid: 6105b6f4-1508-8289-32e2-4dcbbf4dd4d1
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.ProjectGuideUseDefaultContent property (Project)

 **True** if the Project Guide uses the default content. **False** if you want to use custom content for the Project Guide. Read/write **Boolean**.


## Syntax

_expression_. `ProjectGuideUseDefaultContent`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks


> [!NOTE] 
> The Project Guide is deprecated in Project. Instead of the Project Guide, we recommend that you create task pane apps.

However, you can still use custom Project Guides and get the default Project Guide files from the Project SDK download. The Project Guide files are modified for access in a flat folder structure and to remove the  `gbui://` protocol (**gbui** is the goal-based user interface protocol in Office Project 2007 and previous versions). All Project Guide settings must be made programmatically.

The default value of the **ProjectGuideFunctionalLayoutPage** property is `gbui://mainpage.htm`, which does not work because Project does not implement the  `gbui://` protocol. The Project Programmability blog ( `https://blogs.msdn.com/project_programmability/`) includes articles that show how to use the Project Guide in a VBA macro and in an add-in that is developed with Visual C# in Microsoft Office development tools in Visual Studio 2010.


## Example

The following code sample changes the default content for the Project Guide to the XML file specified by the user. An input box prompts the user for the path and file name for custom Project Guide content.




> [!NOTE] 
> Before running this macro, change path to an example path you would like to use, and change filename to the name of an example file, such as custom.xml.




```vb
Sub UseCustomProjectGuide() 
   If Projects.Count = 0 Then 
      MsgBox "You must have at least one active project open." 
      Exit Sub 
   End If 
 
   Dim ProjectGuideURL As String 
   ProjectGuideURL = InputBox$(Prompt:="Enter the path and " _ 
      & "file name of the XML file for custom Project " _ 
      & "Guide content." & Chr(13) _ 
      & "For example, path \filename ") 
   If ProjectGuideURL = Empty Then 
      Exit Sub 
   Else 
      ActiveProject.ProjectGuideUseDefaultContent = False 
      ActiveProject.ProjectGuideContent = ProjectGuideURL 
      MsgBox Prompt:="The custom Project Guide content " _ 
         & "defined in " & ProjectGuideURL & " is " _ 
         & "now in use for the current project." 
   End If 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]