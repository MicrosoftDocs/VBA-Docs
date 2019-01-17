---
title: WizardProperty Object (Publisher)
keywords: vbapb10.chm1638399
f1_keywords:
- vbapb10.chm1638399
ms.prod: publisher
api_name:
- Publisher.WizardProperty
ms.assetid: 9f059422-5454-1902-a092-76e21e36a3f7
ms.date: 06/08/2017
localization_priority: Normal
---


# WizardProperty Object (Publisher)

Represents a setting that is part of a specific publication design or a Design Gallery object's wizard.
 


## Example

Use the  **[Item](Publisher.WizardProperties.Item.md)** property or the **[FindByPropertyID](Publisher.WizardProperties.FindPropertyById.md)** method with the **WizardProperties** collection to return a single **WizardProperty** object. The following example reports on the publication design associated with the active publication, displaying its name and current settings.
 

 

```vb
Dim wizTemp As Wizard 
Dim wizproTemp As WizardProperty 
Dim wizproAll As WizardProperties 
 
Set wizTemp = ActiveDocument.Wizard 
 
With wizTemp 
 Set wizproAll = .Properties 
 Debug.Print "Publication Design associated with " _ 
 &amp; "current publication: " _ 
 &amp; .Name 
 For Each wizproTemp In wizproAll 
 With wizproTemp 
 Debug.Print " Wizard property: " _ 
 &amp; .Name &amp; " = " &amp; .CurrentValueId 
 End With 
 Next wizproTemp 
End With
```


 **Note**  Depending on the language version of Microsoft Publisher that you are using, you may receive an error when using the above code. If this occurs, you will need to build in error handlers to circumvent the errors. For more information, see  **[Wizard Object](Publisher.Wizard.md)**.
 


## Properties



|Name|
|:-----|
|[Application](Publisher.WizardProperty.Application.md)|
|[CurrentValueId](Publisher.WizardProperty.CurrentValueId.md)|
|[Enabled](Publisher.WizardProperty.Enabled.md)|
|[ID](Publisher.WizardProperty.ID.md)|
|[Name](Publisher.WizardProperty.Name.md)|
|[Parent](Publisher.WizardProperty.Parent.md)|
|[Values](Publisher.WizardProperty.Values.md)|

