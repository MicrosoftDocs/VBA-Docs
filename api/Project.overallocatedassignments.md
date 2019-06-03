---
title: OverAllocatedAssignments object (Project)
ms.prod: project-server
ms.assetid: b2856ebf-cff2-04a6-53c9-123de09f2a3b
ms.date: 06/08/2017
localization_priority: Normal
---


# OverAllocatedAssignments object (Project)

Represents a collection of  **[Assignment](Project.Assignment.md)** objects where the resource is overallocated.
 


## Remarks

Use the  **[Item](Project.OverAllocatedAssignments.Item.md)** property to get a single **Assignment** object from the **OverAllocatedAssignments** collection.
 

 

## Example

The following example finds assignments where the resource is overallocated. When the overPeak argument is  **False**, the overallocation is not greater than the maximum resource time available (100%). If you set overPeak to **True**, the example finds overallocated assignments that exceed maximum resource time available, such as 150%.
 

 

```vb
Sub FindOverallocatedAssignments()  
    Dim t As Task  
    Dim a As Assignment  
    Dim overAlloc As OverAllocatedAssignments  
    Dim numOver As Long  
    Dim overPeak As Boolean  
  
    overPeak = False  
  
    For Each t In ActiveProject.Tasks  
        If t.Overallocated Then  
            Set overAlloc = t.StartDriver.OverAllocatedAssignments(overPeak)  
            numOver = overAlloc.Count  
            totalNumOver = overAlloc.TotalDetectedCount  
  
            For Each a In overAlloc  
                Debug.Print "Resource: " & a.Resource.Name & " is overallocated on task: " & t.Name  
                Debug.Print vbTab & "Number of overallocated assignments: " & numOver  
            Next a  
        End If  
    Next t  
End Sub
```


## Properties



|Name|
|:-----|
|[Application](Project.OverAllocatedAssignments.Application.md)|
|[Count](Project.OverAllocatedAssignments.Count.md)|
|[Item](Project.OverAllocatedAssignments.Item.md)|
|[Parent](Project.OverAllocatedAssignments.Parent.md)|
|[TotalDetectedCount](Project.OverAllocatedAssignments.TotalDetectedCount.md)|

## See also


 
[Project Object Model](../project/Concepts/project-object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]