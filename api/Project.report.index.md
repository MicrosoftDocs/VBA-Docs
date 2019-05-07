---
title: Report.Index property (Project)
ms.prod: project-server
ms.assetid: 3a0ccb0f-443e-ea35-4766-b79f97fef84a
ms.date: 06/08/2017
localization_priority: Normal
---


# Report.Index property (Project)
Gets the index of a custom report in the  **Reports** collection. Read-only **Long**.

## Syntax

_expression_.**Index**

_expression_ A variable that represents a 'Report' object.


## Example

The following example lists the index and name of each custom report in a project.


```vb
Sub ListCustomReports()
    Dim oReport As Report
    Dim msg As String
    Dim msgBoxTitle As String
    msg = ""
    msgBoxTitle = "Custom reports in '" & ActiveProject.Name & "'"
    
    For Each oReport In ActiveProject.Reports
        msg = msg & oReport.Index & oReport.Name & vbCrLf
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

 **INT32**


## See also


[Report Object](Project.report.md)
[Reports Object](Project.reports.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]