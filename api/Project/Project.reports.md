---
title: Reports Object (Project)
ms.prod: project-server
ms.assetid: a9f4a13b-1907-dbe8-8077-fb1226bb8bb9
ms.date: 06/08/2017
---


# Reports Object (Project)
Contains a collection of  **[Report](Project.report.md)** objects, where each report is a custom report.
 

## Example

The  **Reports** object is the collection of custom reports in a project. It does not include the built-in reports, such as **Project Overview**,  **Critical Tasks**, or  **Milestone Report**. Use the  **Project.Reports** property to get the **Reports** collection object, as in the following example:
 

 

```
Sub ListCustomReports()
    Dim oReport As Report
    Dim msg As String
    Dim msgBoxTitle As String
    msg = ""
    msgBoxTitle = "Custom reports in '" &amp; ActiveProject.Name &amp; "'"
    
    For Each oReport In ActiveProject.Reports
        msg = msg &amp; oReport.Index &amp; ": " &amp; oReport.Name &amp; vbCrLf
    Next oReport
        
    If ActiveProject.Reports.Count > 0 Then
        MsgBox Prompt:=msg, Title:=msgBoxTitle
    Else
        MsgBox Prompt:="This project contains no custom reports.", _
            Title:=msgBoxTitle
    End If
End Sub
```


## Methods



|**Name**|
|:-----|
|[Add](Project.reports.add.md)|
|[Copy](Project.reports.copy.md)|
|[IsPresent](Project.reports.ispresent.md)|

## Properties



|**Name**|
|:-----|
|[Application](Project.reports.application.md)|
|[Count](Project.reports.count.md)|
|[Item](Project.reports.item.md)|
|[Parent](Project.reports.parent.md)|

## See also


#### Other resources


 
[Report Object](Project.report.md)
 
[Project.Reports Property](project-reports-property-project.md)
