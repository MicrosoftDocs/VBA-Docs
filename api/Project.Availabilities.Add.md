---
title: Availabilities.Add method (Project)
ms.prod: project-server
api_name:
- Project.Availabilities.Add
ms.assetid: 4506674e-947b-905b-93bd-73a58281d676
ms.date: 06/08/2017
localization_priority: Normal
---


# Availabilities.Add method (Project)

Adds an  **Availability** object to an **Availabilities** collection.


## Syntax

_expression_.**Add** (_AvailableFrom_, _AvailableTo_, _AvailableUnit_)

_expression_ A variable that represents an 'Availabilities' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _AvailableFrom_|Required|**Variant**|The earliest date the resource is available for work on the project.|
| _AvailableTo_|Required|**Variant**| The latest date the resource is available for work on the project.|
| _AvailableUnit_|Required|**Double**|The unit value for the availability period.|

## Return value

 **Availability**


## See also


[Availabilities Collection Object](Project.availabilities.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]