---
title: OutlineCode object (Project)
ms.prod: project-server
api_name:
- Project.OutlineCode
ms.assetid: 8f75bdd3-ed5b-ed0f-9c3c-85af3a21580c
ms.date: 06/08/2017
localization_priority: Normal
---


# OutlineCode object (Project)


 

Represents a local outline code in Project. The **OutlineCode** object is a member of the **[OutlineCodes](Project.outlinecodes(object).md)** collection.
 
 **Using the OutlineCode Object**
 
The following example adds a custom outline code to store the location of resources and configures the outline code so that only values specified in the lookup table can be associated with a resource. 
 



```vb
Sub CreateLocationOutlineCode() 
    Dim objOutlineCode As OutlineCode 
 
    Set objOutlineCode = ActiveProject.OutlineCodes.Add( _
        pjCustomResourceOutlineCode1, "Location") 
 
    objOutlineCode.OnlyLookUpTableCodes = True 
 
    DefineLocationCodeMask objOutlineCode.CodeMask 
    EditLocationLookupTable objOutlineCode.LookupTable 
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


## Remarks

An outline code is a type of local custom field that has a hierarchical text lookup table. Enterprise custom fields of type  **Text** that have hierarchical lookup tables act as outline codes. Use the **[OutlineCodes](Project.Project.OutlineCodes.md)** property to return an **OutlineCodes** collection. Use the **[Add](Project.OutlineCodes.Add.md)** method to add a local outline code to the **OutlineCodes** collection. To add an enterprise custom field, you must use Project Web App or the Project Server Interface (PSI).
 

 

## Methods



|Name|
|:-----|
|[Delete](Project.OutlineCode.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.OutlineCode.Application.md)|
|[CodeMask](Project.OutlineCode.CodeMask.md)|
|[DefaultValue](Project.OutlineCode.DefaultValue.md)|
|[FieldID](Project.OutlineCode.FieldID.md)|
|[Index](Project.OutlineCode.Index.md)|
|[LinkedFieldID](Project.OutlineCode.LinkedFieldID.md)|
|[LookupTable](Project.OutlineCode.LookupTable.md)|
|[MatchGeneric](Project.OutlineCode.MatchGeneric.md)|
|[Name](Project.OutlineCode.Name.md)|
|[OnlyCompleteCodes](Project.OutlineCode.OnlyCompleteCodes.md)|
|[OnlyLeaves](Project.OutlineCode.OnlyLeaves.md)|
|[OnlyLookUpTableCodes](Project.OutlineCode.OnlyLookUpTableCodes.md)|
|[Parent](Project.OutlineCode.Parent.md)|
|[RequiredCode](Project.OutlineCode.RequiredCode.md)|
|[SortOrder](Project.OutlineCode.SortOrder.md)|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]