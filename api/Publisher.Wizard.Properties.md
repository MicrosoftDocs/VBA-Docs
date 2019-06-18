---
title: Wizard.Properties property (Publisher)
keywords: vbapb10.chm1441797
f1_keywords:
- vbapb10.chm1441797
ms.prod: publisher
api_name:
- Publisher.Wizard.Properties
ms.assetid: 9f9811b3-10ee-d429-c5a2-8223349525f2
ms.date: 06/18/2019
localization_priority: Normal
---


# Wizard.Properties property (Publisher)

Returns a **[WizardProperties](Publisher.WizardProperties.md)** collection representing all the settings that are part of the specified publication design or Design Gallery object's wizard.


## Syntax

_expression_.**Properties**

_expression_ A variable that represents a **[Wizard](Publisher.Wizard.md)** object.


## Return value

WizardProperties


## Example

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

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]