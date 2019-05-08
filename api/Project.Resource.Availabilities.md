---
title: Resource.Availabilities property (Project)
keywords: vbapj.chm131411
f1_keywords:
- vbapj.chm131411
ms.prod: project-server
api_name:
- Project.Resource.Availabilities
ms.assetid: 1525ba2e-49c1-216a-0b45-008e866163d5
ms.date: 06/08/2017
localization_priority: Normal
---


# Resource.Availabilities property (Project)

Returns an  **[Availabilities](Project.availabilities.md)** collection representing all the available periods defined for the resource in the **Resource Availability** grid. Read-only **Availabilities**.


## Syntax

_expression_. `Availabilities`

_expression_ A variable that represents a [Resource](./Project.Resource.md) object.


## Remarks

The  **Availabilities** property does not return any meaningful information for material resources. Setting a value returns a trappable error (error code 1101) when applied to material resources.


## Example

The following example displays the range of dates during which the specified resource is available for work.


```vb
Sub ShowWorkAvail()
  Dim Avail As Availability
  For Each Avail In ActiveProject.Resources("Tom").Availabilities
    MsgBox "From " & Avail.AvailableFrom & " to " & Avail.AvailableTo
  Next Avail
End Sub
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]