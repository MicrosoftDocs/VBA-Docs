---
title: GroupCriterion2 object (Project)
ms.prod: project-server
api_name:
- Project.GroupCriterion2
ms.assetid: 06047a9d-a9db-43e0-e759-e24560da7128
ms.date: 06/08/2017
localization_priority: Normal
---


# GroupCriterion2 object (Project)

Represents a criterion in a group definition where the group hierarchy can be maintained and cell color can be a hexadecimal value. The  **GroupCriterion2** object is a member of the **[GroupCriteria2](Project.groupcriteria2.md)** collection.
 


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
        Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in ascending order." & vbCrLf  
    Else  
        Fields = Fields & GC2.Index & ". " & GC2.FieldName & " is sorted in descending order." & vbCrLf  
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
|[Delete](Project.GroupCriterion2.Delete.md)|

## Properties



|Name|
|:-----|
|[Application](Project.GroupCriterion2.Application.md)|
|[Ascending](Project.GroupCriterion2.Ascending.md)|
|[Assignment](Project.GroupCriterion2.Assignment.md)|
|[CellColor](Project.GroupCriterion2.CellColor.md)|
|[CellColorEx](Project.GroupCriterion2.CellColorEx.md)|
|[FieldName](Project.GroupCriterion2.FieldName.md)|
|[FontBold](Project.GroupCriterion2.FontBold.md)|
|[FontColor](Project.GroupCriterion2.FontColor.md)|
|[FontColorEx](Project.GroupCriterion2.FontColorEx.md)|
|[FontItalic](Project.GroupCriterion2.FontItalic.md)|
|[FontName](Project.GroupCriterion2.FontName.md)|
|[FontSize](Project.GroupCriterion2.FontSize.md)|
|[FontUnderLine](Project.GroupCriterion2.FontUnderLine.md)|
|[GroupInterval](Project.GroupCriterion2.GroupInterval.md)|
|[GroupOn](Project.GroupCriterion2.GroupOn.md)|
|[Index](Project.GroupCriterion2.Index.md)|
|[Parent](Project.GroupCriterion2.Parent.md)|
|[Pattern](Project.GroupCriterion2.Pattern.md)|
|[StartAt](Project.GroupCriterion2.StartAt.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]