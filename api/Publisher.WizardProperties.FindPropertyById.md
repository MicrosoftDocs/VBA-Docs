---
title: WizardProperties.FindPropertyById method (Publisher)
keywords: vbapb10.chm1507332
f1_keywords:
- vbapb10.chm1507332
ms.prod: publisher
api_name:
- Publisher.WizardProperties.FindPropertyById
ms.assetid: 9d13ffa2-f251-0e7d-2f36-c747413143d0
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardProperties.FindPropertyById method (Publisher)

Returns a **[WizardProperty](Publisher.WizardProperty.md)** object, based on the specified ID, from the collection of wizard properties associated with a publication design or a Design Gallery object's wizard.


## Syntax

_expression_.**FindPropertyById** (_ID_)

_expression_ A variable that represents a **[WizardProperties](Publisher.WizardProperties.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
|_ID_|Required| **Long**|The ID of the wizard property to return; corresponds to the **[ID](Publisher.WizardProperty.ID.md)** property of the **WizardProperty** object.|

## Return value

WizardProperty


## Example

The following example changes the settings of the current publication design (Newsletter Wizard) so that the publication has a region dedicated to the customer's address (Customer Address).

```vb
Sub SetWizardProperties 
 Dim wizTemp As Wizard 
 Dim wizproTemp As WizardProperty 
 
 Set wizTemp = ActiveDocument.Wizard 
 
 With wizTemp.Properties 
 Set wizproTemp = .FindPropertyById(ID:=901) 
 wizproTemp.CurrentValueId = 1 
 End With 
 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]