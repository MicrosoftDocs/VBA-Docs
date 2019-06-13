---
title: Shape.WizardTag property (Publisher)
keywords: vbapb10.chm2228324
f1_keywords:
- vbapb10.chm2228324
ms.prod: publisher
api_name:
- Publisher.Shape.WizardTag
ms.assetid: b93bbdf9-6ce7-3ba6-566a-b11f8044fbda
ms.date: 06/13/2019
localization_priority: Normal
---


# Shape.WizardTag property (Publisher)

Returns or sets a **[PbWizardTag](Publisher.PbWizardTag.md)** constant indicating the function of a specified shape with respect to its publication design. Read/write.


## Syntax

_expression_.**WizardTag**

_expression_ A variable that represents a **[Shape](Publisher.Shape.md)** object.


## Remarks

The **WizardTag** property value can be one of the **PbWizardTag** constants declared in the Microsoft Publisher type library.

The combination of the **[WizardTagInstance](Publisher.Shape.WizardTagInstance.md)** property and the **WizardTag** property uniquely defines every shape in a publication.


## Example

The following example displays the wizard tag and wizard tag instance information for all the shapes on page one of the active publication.

```vb
Dim shpLoop As Shape 
 
For Each shpLoop In ActiveDocument.Pages(1).Shapes 
 With shpLoop 
 Debug.Print "Shape: " & .Name 
 Debug.Print " Wizard tag: " & .WizardTag 
 Debug.Print " Wizard tag instance: " _ 
 & .WizardTagInstance 
 End With 
Next shpLoop
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]