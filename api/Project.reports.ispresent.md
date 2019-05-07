---
title: Reports.IsPresent method (Project)
ms.prod: project-server
ms.assetid: 6040d01a-d187-2f79-945d-1e85b3539a51
ms.date: 06/08/2017
localization_priority: Normal
---


# Reports.IsPresent method (Project)
Indicates whether the specified custom report exists in the project.

## Syntax

_expression_. `IsPresent` _(Name)_

_expression_ A variable that represents a 'Reports' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|Description|

## Return value

 **Boolean**

 **True** if the custom report exists; otherwise, **False**.


## Example

The following example uses the  **IsPresent** method to determine whether a report exists and can be displayed.


```vb
Sub ShowAReport()
    Dim reportName As String
    
    reportName = "Table Tests"
    
    If ActiveProject.Reports.IsPresent(reportName) Then
        ActiveProject.Reports(reportName).Apply
    Else
        MsgBox Prompt:="No custom report name: " & reportName, Title:="Report apply error"
    End If
End Sub
```


## See also


[Reports Object](Project.reports.md)
[Report Object](Project.report.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]