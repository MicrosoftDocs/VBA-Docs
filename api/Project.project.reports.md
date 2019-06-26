---
title: Project.Reports property (Project)
ms.prod: project-server
ms.assetid: dc725fac-a25e-c134-6017-d73060c51e83
ms.date: 06/08/2017
localization_priority: Normal
---


# Project.Reports property (Project)
Gets the collection of custom reports in the project. Read-only  **Reports**.

## Syntax

_expression_. `Reports`

_expression_ A variable that represents a **[Project](project.project.md)** object.


## Example

The  **Reports** object is the collection of custom reports in a project. It does not include the built-in reports, such as **Project Overview**,  **Critical Tasks**, or  **Milestone Report**. Use the  **Project.Reports** property to get the **Reports** collection object, as in the following example:


```vb
Sub ListCustomReports()
    Dim oReport As Report
    Dim msg As String
    Dim msgBoxTitle As String
    msg = ""
    msgBoxTitle = "Custom reports in '" & ActiveProject.Name & "'"
    
    For Each oReport In ActiveProject.Reports
        msg = msg & oReport.Index & ": " & oReport.Name & vbCrLf
    Next oReport
        
    If ActiveProject.Reports.Count > 0 Then
        MsgBox Prompt:=msg, Title:=msgBoxTitle
    Else
        MsgBox Prompt:="This project contains no custom reports.", _
            Title:=msgBoxTitle
    End If
End Sub
```


## Property value

 **REPORTS**


## See also


[Project Object](Project.Project.md)



[Reports Object](Project.reports.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]