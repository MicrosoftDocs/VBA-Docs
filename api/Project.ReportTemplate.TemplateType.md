---
title: ReportTemplate.TemplateType property (Project)
ms.prod: project-server
api_name:
- Project.ReportTemplate.TemplateType
ms.assetid: 5461ae85-0168-f31b-1c04-878afed001e2
ms.date: 06/08/2017
localization_priority: Normal
---


# ReportTemplate.TemplateType property (Project)

Gets the Visual Report template type. Read-only  **PjVisualReportsTemplateType**.


## Syntax

_expression_. `TemplateType`

_expression_ A variable that represents a [ReportTemplate](./Project.ReportTemplate.md) object.


## Remarks

The TemplateType property can be one of the **[PjVisualReportsTemplateType](Project.PjVisualReportsTemplateType.md)** constants.


## Example

The following example lists all of the Visual Report template types and files for the current user.


```vb
Sub ListTemplatePaths() 

 Dim templateList As String 

 Dim typeOfTemplate As String 

 Dim template As ReportTemplate 

 

 For Each template In Application.VisualReportTemplateList 

 Select Case template.TemplateType 

 Case pjExcel 

 typeOfTemplate = "Excel" 

 Case pjVisioMetric 

 typeOfTemplate = "Visio Metric" 

 Case pjVisioUS 

 typeOfTemplate = "Visio U.S." 

 Case Else 

 End Select 

 

 templateList = templateList & vbCrLf & typeOfTemplate & ": " _ 

 & template.TemplatePath 

 Next template 

 

 MsgBox "Visual Reports Templates:" & templateList 

 

End Sub
```


## See also


[ReportTemplate Object](Project.ReportTemplate.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]