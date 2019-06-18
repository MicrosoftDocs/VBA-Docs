---
title: WizardProperty.Values property (Publisher)
keywords: vbapb10.chm1572872
f1_keywords:
- vbapb10.chm1572872
ms.prod: publisher
api_name:
- Publisher.WizardProperty.Values
ms.assetid: 478d3b98-65f4-c448-8096-3e999c865846
ms.date: 06/18/2019
localization_priority: Normal
---


# WizardProperty.Values property (Publisher)

Returns a **[WizardValues](Publisher.WizardValues.md)** collection representing all the valid values for a wizard property.


## Syntax

_expression_.**Values**

_expression_ A variable that represents a **[WizardProperty](Publisher.WizardProperty.md)** object.


## Return value

WizardValues


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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]