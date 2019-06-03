---
title: WizardValue object (Publisher)
keywords: vbapb10.chm2162687
f1_keywords:
- vbapb10.chm2162687
ms.prod: publisher
api_name:
- Publisher.WizardValue
ms.assetid: 15b60632-d1b1-c62b-0264-72d65bd1fe82
ms.date: 06/04/2019
localization_priority: Normal
---


# WizardValue object (Publisher)

Represents a possible value for the specified wizard property.
 
## Remarks

Use the **[Item](Publisher.WizardValues.Item.md)** property of the **WizardValues** collection to return a **WizardValue** object. 

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

- [Application](Publisher.WizardValue.Application.md)
- [ID](Publisher.WizardValue.ID.md)
- [Name](Publisher.WizardValue.Name.md)
- [Parent](Publisher.WizardValue.Parent.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]