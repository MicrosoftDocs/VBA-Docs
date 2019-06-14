---
title: Page.Wizard property (Publisher)
keywords: vbapb10.chm393234
f1_keywords:
- vbapb10.chm393234
ms.prod: publisher
api_name:
- Publisher.Page.Wizard
ms.assetid: 05cf1482-bde5-9ea2-4099-69a56a2dc61a
ms.date: 06/11/2019
localization_priority: Normal
---


# Page.Wizard property (Publisher)

Returns a **[Wizard](Publisher.Wizard.md)** object representing the publication design associated with the specified publication or the wizard associated with the specified Design Gallery object.


## Syntax

_expression_.**Wizard**

_expression_ A variable that represents a **[Page](Publisher.Page.md)** object.


## Remarks

When accessing the **Wizard** property from the **Document** or **Page** object, if the specified publication is not associated with any publication design, an error occurs. 

When accessing the **Wizard** property from the **Shape** or **ShapeRange** object, if the specified object is not a Design Gallery object, an error occurs.


## Example

The following example reports on the publication design associated with the active publication, displaying its name and current settings.

```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication design associated with " _ 
 & "current publication: " _ 
 & .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Setting: " _ 
 & .Name & " = " & .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```

> [!NOTE] 
> Depending on the language version of Publisher that you are using, you may receive an error when using this code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see the **[Wizard](Publisher.Wizard.md)** object.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]