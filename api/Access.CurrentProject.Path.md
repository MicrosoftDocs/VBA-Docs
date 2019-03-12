---
title: CurrentProject.Path property (Access)
keywords: vbaac10.chm12718
f1_keywords:
- vbaac10.chm12718
ms.prod: access
api_name:
- Access.CurrentProject.Path
ms.assetid: 25f28502-b5fc-aafa-9189-eb091907a529
ms.date: 02/27/2019
localization_priority: Normal
---


# CurrentProject.Path property (Access)

You can use the **Path** property to determine the location where data is stored for a Microsoft Access project (.adp) or Microsoft Access database. Read-only **String**.


## Syntax

_expression_.**Path**

_expression_ A variable that represents a **[CurrentProject](Access.CurrentProject.md)** object.


## Remarks

You can use the **Path** property to determine the location of information stored through the **[CurrentProject](Access.CurrentProject.md)** or **[CodeProject](Access.CodeProject.md)** objects of a project or database.


## Example

The following example displays a message indicating the disk location of the current Access project or database.


```vb
MsgBox "The current database is located at " & Application.CurrentProject.Path & "." 
 
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
