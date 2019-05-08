---
title: Reports.Add method (Project)
ms.prod: project-server
ms.assetid: 3ce8e51c-54c6-6cc7-f5ec-c27e0a657f04
ms.date: 06/08/2017
localization_priority: Normal
---


# Reports.Add method (Project)
Adds a custom report to the  **Reports** collection.

## Syntax

_expression_.**Add** _(Name)_

_expression_ A variable that represents a 'Reports' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the report.|
| _Name_|Required|**String**||

## Return value

 **Report**

The custom report object that is added.


## Remarks

The new report is empty; it does not contain any shapes such as tables or charts. To add shapes to the report, you can use methods in the  **[Shapes](Project.shapes.md)** object such as **AddChart** and **AddTable**.


## Examples

The following example adds an empty report named  **Report 1**, and displays the report.


```vb
Sub AddReport()
    ActiveProject.Reports.Add "Report 1"
End Sub
```

To delete a report, you must change to a different view, as in the following example:




```vb
Sub DeleteAReport()
    Dim reportName As String
    
    reportName = "Report 1"
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ' To delete the active report, change to another view.
        ViewApplyEx Name:="&Gantt Chart"
        
        ActiveProject.Reports(reportName).Delete
    Else
        MsgBox Prompt:="No report name: " & reportName, Title:="Report delete error"
    End If
End Sub
```


## See also


[Reports Object](Project.reports.md)
[Report Object](Project.report.md)
[Shapes](Project.shapes.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]