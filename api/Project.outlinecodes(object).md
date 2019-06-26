---
title: OutlineCodes object (Project)
ms.prod: project-server
ms.assetid: a2e6d0c7-0741-91c6-61aa-f4bcc299e66f
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlineCodes object (Project)

Contains a collection of  **[OutlineCode](Project.OutlineCode.md)** objects.
 


## Remarks

An outline code is a type of local custom field that has a hierarchical text lookup table. Enterprise custom fields of type  **Text** that have hierarchical lookup tables act as outline codes. Use the **[OutlineCodes](Project.Project.OutlineCodes.md)** property to return an **OutlineCodes** collection. Use the **[Add](Project.OutlineCodes.Add.md)** method to add a local outline code to the **OutlineCodes** collection. To add an enterprise custom field, you must use Project Web App or the Project Server Interface (PSI).
 

 

## Example

 **Using the OutlineCodes Collection Object**
 

 
The following example adds a custom outline code to store the location of resources and configures the outline code such that only values specified in the lookup table can be associated with a resource. 
 

 

> [!NOTE] 
> The  **OnlyLookUpTableCodes** property can be set only after the lookup table contains entries. If you try to set **OnlyLookUpTableCodes** before creating lookup table entries, the result is run-time error 7, "Out of memory."
 




```vb
Sub CreateLocationOutlineCode() 

 

 Dim objOutlineCode As OutlineCode 

 

 Set objOutlineCode = ActiveProject.OutlineCodes.Add( _ 

 pjCustomResourceOutlineCode1, "Location") 

 

 DefineLocationCodeMask objOutlineCode.CodeMask 

 EditLocationLookupTable objOutlineCode.LookupTable 

 

 objOutlineCode.OnlyLookUpTableCodes = True 

 

End Sub 

 

 

Sub DefineLocationCodeMask(objCodeMask As CodeMask) 

 objCodeMask.Add _ 

 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 

 Length:=2, Separator:="." 

 

 objCodeMask.Add _ 

 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 

 Separator:="." 

 

 objCodeMask.Add _ 

 Sequence:=pjCustomOutlineCodeUppercaseLetters, _ 

 Length:=3, Separator:="." 

End Sub 

 

 

Sub EditLocationLookupTable(objLookupTable As LookupTable) 

 Dim objStateEntry As LookupTableEntry 

 Dim objCountyEntry As LookupTableEntry 

 Dim objCityEntry As LookupTableEntry 

 

 Set objStateEntry = objLookupTable.AddChild("WA") 

 objStateEntry.Description = "Washington" 

 

 Set objCountyEntry = objLookupTable.AddChild("KING", _ 

 objStateEntry.UniqueID) 

 objCountyEntry.Description = "King County" 

 

 Set objCityEntry = objLookupTable.AddChild("SEA", _ 

 objCountyEntry.UniqueID) 

 objCityEntry.Description = "Seattle" 

 

 Set objCityEntry = objLookupTable.AddChild("RED", _ 

 objCountyEntry.UniqueID) 

 objCityEntry.Description = "Redmond" 

 

 Set objCityEntry = objLookupTable.AddChild("KIR", _ 

 objCountyEntry.UniqueID) 

 objCityEntry.Description = "Kirkland" 

End Sub
```


## Methods



|Name|
|:-----|
|[Add](Project.OutlineCodes.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Project.OutlineCodes.Application.md)|
|[Count](Project.OutlineCodes.Count.md)|
|[Item](Project.OutlineCodes.Item.md)|
|[Parent](Project.OutlineCodes.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]