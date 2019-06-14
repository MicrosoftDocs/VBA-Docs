---
title: WizardProperty object (Publisher)
keywords: vbapb10.chm1638399
f1_keywords:
- vbapb10.chm1638399
ms.prod: publisher
api_name:
- Publisher.WizardProperty
ms.assetid: 9f059422-5454-1902-a092-76e21e36a3f7
ms.date: 06/04/2019
localization_priority: Normal
---


# WizardProperty object (Publisher)

Represents a setting that is part of a specific publication design or a Design Gallery object's wizard.
 
## Remarks

Use the **[Item](Publisher.WizardProperties.Item.md)** property or the **[FindByPropertyID](Publisher.WizardProperties.FindPropertyById.md)** method of the **WizardProperties** collection to return a single **WizardProperty** object. 

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
 Debug.Print "Publication Design associated with " _ 
 & "current publication: " _ 
 & .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 & .Name & " = " & .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


## Properties

- [Application](Publisher.WizardProperty.Application.md)
- [CurrentValueId](Publisher.WizardProperty.CurrentValueId.md)
- [Enabled](Publisher.WizardProperty.Enabled.md)
- [ID](Publisher.WizardProperty.ID.md)
- [Name](Publisher.WizardProperty.Name.md)
- [Parent](Publisher.WizardProperty.Parent.md)
- [Values](Publisher.WizardProperty.Values.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]