---
title: Project.Type property (Project)
ms.prod: project-server
api_name:
- Project.Project.Type
ms.assetid: 13393b8e-283d-d816-283e-f363b83eac91
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Type property (Project)

Gets the type of a project. Read-only  **PjProjectType**.


## Syntax

_expression_.**Type**

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Remarks

The **Type** property can be one of the **[PjProjectType](Project.PjProjectType.md)** constants.


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


## See also


[Project Object](Project.Project.md)
[PjProjectType Enumeration](Project.PjProjectType.md)



[CheckoutProject Method](Project.project.checkoutproject.md)
[Application.IsCheckedOut Property](Project.application.ischeckedout.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]