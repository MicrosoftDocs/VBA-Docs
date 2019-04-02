---
title: Resource.HyperlinkAddress property (Project)
ms.prod: project-server
api_name:
- Project.Resource.HyperlinkAddress
ms.assetid: 44de3c24-ff9d-49dc-d942-8167a73b9ef6
ms.date: 03/05/2019
localization_priority: Normal
---


# Resource.HyperlinkAddress property (Project)

Gets or sets the URL or UNC path of a document. Read/write **String**.


## Syntax

_expression_.**HyperlinkAddress**

_expression_ A variable that represents a **[Resource](Project.Resource.md)** object.


## Example

The following example adds a hyperlink to all tasks in the active project, including tasks in subprojects.

```vb
Sub AddHyperlink() 
 Dim T As Task 
 
 For Each T In ActiveProject.Tasks 
 If Not (T Is Nothing) Then 
 T.Hyperlink = "Microsoft" 
 T.HyperlinkAddress = "https://www.microsoft.com/" 
 End If 
 Next T 
 
End Su
```


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]