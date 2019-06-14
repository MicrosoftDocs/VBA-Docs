---
title: Wizard object (Publisher)
keywords: vbapb10.chm1507327
f1_keywords:
- vbapb10.chm1507327
ms.prod: publisher
api_name:
- Publisher.Wizard
ms.assetid: c0a64ee9-d1fa-6dc7-5221-ff2d32874ea0
ms.date: 06/04/2019
localization_priority: Normal
---


# Wizard object (Publisher)

Represents the publication design associated with a publication, or the wizard associated with a Design Gallery object.
 
## Remarks

Use the **[Wizard](Publisher.Document.Wizard.md)** property of a **Document**, **Page**, **Shape**, or **ShapeRange** object to return a **Wizard** object. 

## Example

> [!NOTE] 
> Depending on the language version of Publisher that you are using, you may receive an error when using this code. If this occurs, you will need to build in error handlers to circumvent the errors. 

This example reports on the publication design associated with the active publication, displaying its name and current settings.

```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 MsgBox "Publication Design associated with " _ 
 & "current publication: " _ 
 & .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 MsgBox " Wizard property: " _ 
 & .Name & " = " & .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```

<br/>

The following example functions as the previous code, but has error handlers built in for this situation.

```vb
Sub ExampleWithErrorHandlers() 
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
 If wizproTemp.Name = "Layout" Or wizproTemp _ 
 .Name = "Layout (Intl)" Then 
 On Error GoTo Handler 
 MsgBox " Wizard property: " _ 
 & .Name & " = " & .CurrentValueId 
 
Handler: 
 If Err.Number = 70 Then Resume Next 
 Else 
 MsgBox " Wizard property: " _ 
 & .Name & " = " & .CurrentValueId 
 End If 
 End With 
 Next wizproTemp 
 End With 
End Sub
```


## Methods

- [SetId](Publisher.Wizard.SetId.md)

## Properties

- [Application](Publisher.Wizard.Application.md)
- [ID](Publisher.Wizard.ID.md)
- [Name](Publisher.Wizard.Name.md)
- [Parent](Publisher.Wizard.Parent.md)
- [Properties](Publisher.Wizard.Properties.md)

## See also

- [Publisher Object Model Reference](overview/publisher/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]