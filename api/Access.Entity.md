---
title: Entity object (Access)
keywords: vbaac10.chm14565
f1_keywords:
- vbaac10.chm14565
ms.prod: access
api_name:
- Access.Entity
ms.assetid: fbce3ef6-bca4-92c6-c191-fd89ad33e888
ms.date: 06/08/2017
localization_priority: Normal
---


# Entity object (Access)

Represents an entity defined in a Data Service data connection.


## Remarks

Use the  **[Item](Access.Entities.Item.md)** property of the **[Entities](Access.Entities.md)** to return an **[Entity](Access.Entity.md)** object.

Use the  **[Operations](Access.Operations.md)** property to returnt he operations defined for the specified entity.

A Data Service data connection may contain one or more entities. Each entity specifies an external content type. Used throughout the functionality and services offered by Business Connectivity Services, external content types are reusable metadata descriptions of connectivity information and data definitions plus the behaviors you want to apply to a certain category of external data. 

For more information about external content types, see [What Are External Content Types?](https://msdn.microsoft.com/library/ee556391%28office.14%29.aspx).


## Properties



|Name|
|:-----|
|[Name](Access.Entity.Name.md)|
|[Operations](Access.Entity.Operations.md)|
|[Parent](Access.Entity.Parent.md)|

## See also


[Access Object Model Reference](overview/Access/object-model.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]