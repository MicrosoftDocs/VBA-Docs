---
title: WizardValues Object (Publisher)
keywords: vbapb10.chm1703935
f1_keywords:
- vbapb10.chm1703935
ms.prod: publisher
api_name:
- Publisher.WizardValues
ms.assetid: 559659bb-6c9f-9325-c931-14044c059e18
ms.date: 06/08/2017
localization_priority: Normal
---


# WizardValues Object (Publisher)

Represents the complete set of valid values for a wizard property.
 


## Example

Use the  **[Values](Publisher.WizardProperty.Values.md)** property of the **WizardProperty** object to return a **WizardValues** collection. The following example displays the current value for the first wizard property in the active publication and then lists all the other possible values.
 

 

```vb
Dim valAll As WizardValues 
Dim valLoop As WizardValue 
 
With ActiveDocument.Wizard 
 Set valAll = .Properties(1).Values 
 
 MsgBox "Wizard: " &amp; .Name &amp; vbLf &amp; _ 
 "Property: " &amp; .Properties(1).Name &amp; vbLf &amp; _ 
 "Current value: " &amp; .Properties(1).CurrentValueId 
 
 For Each valLoop In valAll 
 MsgBox "Possible value: " &amp; valLoop.ID &amp; " (" &amp; valLoop.Name &amp; ")" 
 Next valLoop 
End With
```


## Properties



|Name|
|:-----|
|[Application](Publisher.WizardValues.Application.md)|
|[Count](Publisher.WizardValues.Count.md)|
|[Item](Publisher.WizardValues.Item.md)|
|[Parent](Publisher.WizardValues.Parent.md)|

