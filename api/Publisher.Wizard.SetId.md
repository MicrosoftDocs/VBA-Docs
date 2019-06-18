---
title: Wizard.SetId method (Publisher)
keywords: vbapb10.chm1441798
f1_keywords:
- vbapb10.chm1441798
ms.prod: publisher
api_name:
- Publisher.Wizard.SetId
ms.assetid: d2278716-514b-0858-d68e-868d0daf86b0
ms.date: 06/18/2019
localization_priority: Normal
---


# Wizard.SetId method (Publisher)

Specifies the type of the wizard (template) to which to convert the current publication type.


## Syntax

_expression_.**SetId** (_ID_)

_expression_ A variable that represents a **[Wizard](Publisher.Wizard.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ID_|Required| **Long**|ID of the wizard (template) to which to convert the current publication.|

## Remarks

When Microsoft Publisher converts the publication type, it automatically maps elements of the existing publication type to the new publication type as best as possible. Any elements that it cannot map appear under **Extra Content** in the **Format Publication** task pane in the Publisher user interface; you can add them to the new publication manually by dragging them onto the publication page.


## Example

The following Microsoft Visual Basic for Applications (VBA) macro shows how to use the **SetId** method to convert the current publication type to a brochure.

```vb
Public Sub SetId_Example() 
 
 Wizard.SetId (pbWizardBrochures) 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]