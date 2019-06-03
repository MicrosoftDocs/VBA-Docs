---
title: WizardValues object (Publisher)
keywords: vbapb10.chm1703935
f1_keywords:
- vbapb10.chm1703935
ms.prod: publisher
api_name:
- Publisher.WizardValues
ms.assetid: 559659bb-6c9f-9325-c931-14044c059e18
ms.date: 06/04/2019
localization_priority: Normal
---


# WizardValues object (Publisher)

Represents the complete set of valid values for a wizard property.
 
## Remarks

Use the **[Values](Publisher.WizardProperty.Values.md)** property of the **WizardProperty** object to return a **WizardValues** collection. 

## Example

The following example displays the current value for the first wizard property in the active publication and then lists all the other possible values.

```vb
Dim valAll As WizardValues 
Dim valLoop As WizardValue 
 
With ActiveDocument.Wizard 
 Set valAll = .Properties(1).Values 
 
 MsgBox "Wizard: " & .Name & vbLf & _ 
 "Property: " & .Properties(1).Name & vbLf & _ 
 "Current value: " & .Properties(1).CurrentValueId 
 
 For Each valLoop In valAll 
 MsgBox "Possible value: " & valLoop.ID & " (" & valLoop.Name & ")" 
 Next valLoop 
End With
```


## Properties

- [Application](Publisher.WizardValues.Application.md)
- [Count](Publisher.WizardValues.Count.md)
- [Item](Publisher.WizardValues.Item.md)
- [Parent](Publisher.WizardValues.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]