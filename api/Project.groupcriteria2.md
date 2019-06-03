---
title: GroupCriteria2 object (Project)
ms.prod: project-server
ms.assetid: ac785cc4-dbe3-0b1d-d1f1-6d45c93bfb1d
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupCriteria2 object (Project)

Contains a collection of  **[GroupCriterion2](Project.GroupCriterion2.md)** objects, where the group hierarchy can be maintained and cell color can be a hexadecimal value.
 


## Example

 **Using the GroupCriterion2 Object**
 

 
Use  **GroupCriteria2(***Index* **)**, where*Index* is the criterion index, to return a single **GroupCriterion2** object. The following example sets the cell color for the first criterion in the Standard Rate resource group to blue.
 

 



```vb
ActiveProject.ResourceGroups2("Standard Rate").GroupCriteria2(1).CellColor = &HFF0000
```

 **Using the GroupCriteria2 Collection**
 

 
Use the  **[GroupCriteria](Project.Group2.GroupCriteria.md)** property to return a **GroupCriteria2** collection. The following example displays a list of the fields used as criteria in the specified task group and shows whether they are sorted in ascending or descending order.
 

 



```vb
Dim GC2 As GroupCriterion2  
Dim Fields As String  
  
For Each GC2 In ActiveProject.TaskGroups2("Priority Keeping Outline Structure").GroupCriteria  
    If GC2.Ascending = True Then  
       Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in ascending order." _
           & vbCrLf  
    Else  
        Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in descending order." _
           & vbCrLf  
    End If  
Next GC2  

MsgBox Fields
```

Use the  **[AddEx](Project.GroupCriteria2.AddEx.md)** method to add a **GroupCriterion2** object to the **GroupCriteria2** collection, where **CellColor** can be a hexadecimal value. The following example adds another criterion to the specified resource group, grouping resources in ascending order as determined by the percentage of their work (in 25-percent increments) that is complete.
 

 



```vb
ActiveProject.ResourceGroups2("Response Pending").GroupCriteria2.AddEx "% Work Complete", True, _  
    CellColor:=&H0101FF, GroupOn:=pjGroupOnPct1_25
```


## Methods



|Name|
|:-----|
|[Add](Project.GroupCriteria2.Add.md)|
|[AddEx](Project.GroupCriteria2.AddEx.md)|

## Properties



|Name|
|:-----|
|[Application](Project.GroupCriteria2.Application.md)|
|[Count](Project.GroupCriteria2.Count.md)|
|[Item](Project.GroupCriteria2.Item.md)|
|[Parent](Project.GroupCriteria2.Parent.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]