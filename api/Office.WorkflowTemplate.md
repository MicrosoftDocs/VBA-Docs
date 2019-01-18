---
title: WorkflowTemplate object (Office)
keywords: vbaof11.chm282000
f1_keywords:
- vbaof11.chm282000
ms.prod: office
api_name:
- Office.WorkflowTemplate
ms.assetid: 965d0474-dd51-9b0e-b34c-a11f921ff410
ms.date: 06/08/2017
localization_priority: Normal
---


# WorkflowTemplate object (Office)

Represents one of the workflows available for the current document.


## Remarks

A  **WorkflowTemplate** object corresponds to one of the options displayed in the **Start New Workflow** dialog box. On a Web page, the workflow templates are displayed as a list of options.


## Example

The following example displays the name of each workflow template in the current document and then displays workflow specific configuration user interface for a specific template. It should be noted that calling the  **GetWorkflowTemplates** method involves a round-trip to the server.


```vb
Sub DisplayWorkTemplates() 
Dim objWorkflowTemplates As WorkflowTemplates 
Dim objWorkflowTemplate As WorkflowTemplate 
Dim cnt As Integer 
 
Set objWorkflowTemplates = Document.GetWorkflowTemplates() 
 
For cnt = 1 To objWorkflowTemplates.Count 
 Debug.Print objWorkflowTemplate(cnt).Name 
Next 
 
Set objWorkflowTemplate = objWorkflowTemplates(1) 
objWorkflowTemplate.Show 
 
End Sub 

```


## Methods



|Name|
|:-----|
|[Show](Office.WorkflowTemplate.Show.md)|

## Properties



|Name|
|:-----|
|[Application](Office.WorkflowTemplate.Application.md)|
|[Creator](Office.WorkflowTemplate.Creator.md)|
|[Description](Office.WorkflowTemplate.Description.md)|
|[DocumentLibraryName](Office.WorkflowTemplate.DocumentLibraryName.md)|
|[DocumentLibraryURL](Office.WorkflowTemplate.DocumentLibraryURL.md)|
|[Id](Office.WorkflowTemplate.Id.md)|
|[Name](Office.WorkflowTemplate.Name.md)|

## See also





[Object Model Reference](./overview/Library-Reference/reference-object-library-reference-for-office.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]