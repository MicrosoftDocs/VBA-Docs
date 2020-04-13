---
title: Application.IsCheckedOut property (Project)
ms.prod: project-server
ms.assetid: 616f9342-9d9b-dd85-873c-3e40abfec019
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.IsCheckedOut property (Project)
Gets whether an open project is checked out from Project Web App by the user. Read-only  **Boolean**.

## Syntax

_expression_. `IsCheckedOut`

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ProjectName_|Required|**String**|The name of a project that is open in Project Professional.|

## Remarks

For a project that is open in Project Professional, the **IsCheckedOut** property value is **True** if the project is checked out by the current user. If the specified project is not checked out by the current user (that is, the project is open but in a read-only mode), or is checked out by a different user, the **IsCheckedOut** value is **False**.

The **IsCheckedOut** property returns run-time error 1004, "An unexpected error occurred with the method" in the following cases:


- The specified project is not open in Project Professional.
    
- The specified project is a local project file such as Project1.mpp.
    

## Example

The following example determines whether an open project is an enterprise project and is checked out. If the project is not checked out, the example tries to check out the project. If the project is checked out by another user, Project shows a dialog box with the message, "To check out, DOMAIN\UserName must close the project in their session or contact your administrator to check in the project."


```vb
Sub CheckOutOpenEnterpriseProjects()
    Dim openProjects As Projects
    Dim proj As Project
    
    Set openProjects = Application.Projects
    
    On Error Resume Next
    
    For Each proj In openProjects
        If Application.IsCheckedOut(proj.Name) Then
            If proj.Type = pjProjectTypeEnterpriseCheckedOut Then
                Debug.Print "'" & proj.Name & "'" & " is already checked out."
            ElseIf proj.Type = pjProjectTypeNonEnterprise Then
                Debug.Print "'" & proj.Name & "'" & " is not an enterprise project."
            End If
        Else
            proj.CheckoutProject
            Debug.Print "Attempted to check out: '" & proj.Name & "'"
        End If
    Next proj
End Sub
```


## Property value

 **BOOL**


## See also


[Application Object](Project.Application.md)
[Project.Type Property](Project.Project.Type.md)



[Project.CheckoutProject Method](Project.project.checkoutproject.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]