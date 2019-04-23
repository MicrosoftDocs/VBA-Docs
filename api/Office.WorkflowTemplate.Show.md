---
title: WorkflowTemplate.Show method (Office)
keywords: vbaof11.chm282006
f1_keywords:
- vbaof11.chm282006
ms.prod: office
api_name:
- Office.WorkflowTemplate.Show
ms.assetid: aa4780b5-f3bd-431f-8cb3-20c6058ebc5a
ms.date: 01/29/2019
localization_priority: Normal
---


# WorkflowTemplate.Show method (Office)

Displays a workflow-specific configuration user interface for the specified **WorkflowTemplate** object.


## Syntax

_expression_.**Show**

_expression_ An expression that returns a **[WorkflowTemplate](Office.WorkflowTemplate.md)** object.


## Return value

Integer


## Example

The following example displays the name of each workflow template in the current document and then displays a workflow-specific configuration user interface for a specific template.


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

- [WorkflowTemplate object members](overview/Library-Reference/workflowtemplate-members-office.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]