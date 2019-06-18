---
title: WizardProperty.Enabled property (Publisher)
keywords: vbapb10.chm1572871
f1_keywords:
- vbapb10.chm1572871
ms.prod: publisher
api_name:
- Publisher.WizardProperty.Enabled
ms.assetid: c66741c8-1493-ac90-4ecb-ed8d58743c69
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardProperty.Enabled property (Publisher)

**True** if a wizard property is enabled. Read-only **Boolean**.


## Syntax

_expression_.**Enabled**

_expression_ A variable that represents a **[WizardProperty](Publisher.WizardProperty.md)** object.


## Return value

Boolean


## Example

This example displays the name of each enabled wizard property in the active publication.

```vb
Sub SetEnabledProperty() 
 Dim wizProperty As WizardProperty 
 For Each wizProperty In ActiveDocument.Wizard.Properties 
 If wizProperty.Enabled = True Then 
 MsgBox "The name of the wizard property is " & wizProperty.Name 
 End If 
 Next 
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]