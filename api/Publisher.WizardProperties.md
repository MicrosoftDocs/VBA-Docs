---
title: WizardProperties object (Publisher)
keywords: vbapb10.chm1572863
f1_keywords:
- vbapb10.chm1572863
ms.prod: publisher
api_name:
- Publisher.WizardProperties
ms.assetid: b3feecf2-ffbb-79de-8586-6a64df1b816a
ms.date: 06/04/2019
localization_priority: Normal
---


# WizardProperties object (Publisher)

Represents the settings available in a publication design or in a Design Gallery object's wizard.
 
## Remarks

Use the **[Properties](Publisher.Wizard.Properties.md)** property of a **Wizard** object to return a **WizardProperties** collection. 

## Example

> [!NOTE] 
> Depending on the language version of Publisher that you are using, you may receive an error when using this code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see the **[Wizard](Publisher.Wizard.md)** object.

The following example reports on the publication design associated with the active publication, displaying its name and current settings.

```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 & "current publication: " .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 & .Name & " = " & .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


## Methods

- [FindPropertyById](Publisher.WizardProperties.FindPropertyById.md)

## Properties

- [Application](Publisher.WizardProperties.Application.md)
- [Count](Publisher.WizardProperties.Count.md)
- [Item](Publisher.WizardProperties.Item.md)
- [Parent](Publisher.WizardProperties.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]