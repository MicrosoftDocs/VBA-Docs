---
title: AccessObject.DateModified property (Access)
keywords: vbaac10.chm12753
f1_keywords:
- vbaac10.chm12753
ms.prod: access
api_name:
- Access.AccessObject.DateModified
ms.assetid: a5392776-febe-de09-103d-2d2683f2d0bf
ms.date: 02/01/2019
localization_priority: Normal
---


# AccessObject.DateModified property (Access)

Returns a **Date** indicating the date and time when the design of the specified object was last modified. Read-only.


## Syntax

_expression_.**DateModified**

_expression_ A variable that represents an **[AccessObject](Access.AccessObject.md)** object.


## Example

The following example lists all the reports in the current database and when their designs were created and modified.


```vb
Dim acobjLoop As AccessObject 
 
For Each acobjLoop In CurrentProject.AllReports 
 With acobjLoop 
 Debug.Print .Name & " - Created " & .DateCreated _ 
 & " - Modified " & .DateModified 
 End With 
Next acobjLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]