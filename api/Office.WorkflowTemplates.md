---
title: WorkflowTemplates object (Office)
keywords: vbaof11.chm283000
f1_keywords:
- vbaof11.chm283000
ms.prod: office
api_name:
- Office.WorkflowTemplates
ms.assetid: 01df4716-4440-7761-8504-22f78e40f8e4
ms.date: 01/29/2019
localization_priority: Normal
---


# WorkflowTemplates object (Office)

Represents a collection of **[WorkflowTemplate](office.workflowtemplate.md)** objects.


## Example

The following example displays the name of each workflow template in the current document and then displays a workflow-specific configuration user interface for a specific template. It should be noted that calling the **GetWorkflowTemplates** method involves a round-trip to the server.


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


## See also

- [WorkflowTemplates object members](overview/Library-Reference/workflowtemplates-members-office.md)
- [Object Model Reference](overview/Library-Reference/reference-object-library-reference-for-office.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]