---
title: GroupCriteria object (Project)
ms.prod: project-server
ms.assetid: b19beefb-bfe2-54ba-0835-11624e92bafc
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupCriteria object (Project)

Contains a collection of  **[GroupCriterion](Project.GroupCriterion.md)** objects.
 


## Remarks

For groups where the group hierarchy can be maintained and cell color can be a hexadecimal value, use the **[GroupCriteria2](Project.groupcriteria2.md)** collection object.
 

 

## Example

 **Using the GroupCriterion Object**
 

 
Use  **GroupCriteria(***Index* **)**, where*Index* is the criterion index, to return a single **GroupCriterion** object. The following example sets the cell color for the first criterion in the Standard Rate resource group to blue.
 

 



```vb
ActiveProject.ResourceGroups("Standard Rate").GroupCriteria(1).CellColor = pjBlue
```

 **Using the GroupCriteria Collection**
 

 
Use the **[GroupCriteria](Project.Group.GroupCriteria.md)** property to return a **GroupCriteria** collection. The following example displays a list of the fields used as criteria in the specified task group and whether they are sorted in ascending or descending order.
 

 



```vb
Dim GC As GroupCriterion 

Dim Fields As String 

 

For Each GC In ActiveProject.TaskGroups("Priority Keeping Outline Structure").GroupCriteria 

 If GC.Ascending = True Then 

 Fields = Fields & GC.Index & ". " & GC.FieldName & " is sorted in ascending order." & vbCrLf 

 Else 

 Fields = Fields & GC.Index & ". " & GC.FieldName & " is sorted in descending order." & vbCrLf 

 End If 

Next GC 

 

MsgBox Fields
```

Use the **[Add](Project.GroupCriteria.Add.md)** method to add a **GroupCriterion** object to the **GroupCriteria** collection. The following example adds another criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work (in 25-percent increments) that is complete.
 

 



```vb
ActiveProject.ResourceGroups("Response Pending").GroupCriteria.Add "% Work Complete", True, CellColor:=pjRed, GroupOn:=pjGroupOnPct1_25
```


## Methods



|Name|
|:-----|
|[Add](Project.GroupCriteria.Add.md)|

## Properties



|Name|
|:-----|
|[Application](Project.GroupCriteria.Application.md)|
|[Count](Project.GroupCriteria.Count.md)|
|[Item](Project.GroupCriteria.Item.md)|
|[Parent](Project.GroupCriteria.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]