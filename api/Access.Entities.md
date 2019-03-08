---
title: Entities object (Access)
keywords: vbaac10.chm14560
f1_keywords:
- vbaac10.chm14560
ms.prod: access
api_name:
- Access.Entities
ms.assetid: 8d91418d-ab38-77b1-e767-250b0eb57cb1
ms.date: 03/08/2019
localization_priority: Normal
---


# Entities object (Access)

Represents the collection of entities defined in a Data Service data connection.


## Remarks

A Data Service data connection may contain one or more entities. Each entity specifies an external content type. Used throughout the functionality and services offered by Business Connectivity Services, external content types are reusable metadata descriptions of connectivity information and data definitions plus the behaviors that you want to apply to a certain category of external data. 

Use the **[Entities](Access.WebService.Entities.md)** property to return the entities defined for a Data Service data connection.

Use the **Item** property to return an **[Entity](Access.Entity.md)** object.

For more information about external content types, see [What Are External Content Types](https://docs.microsoft.com/previous-versions/office/developer/sharepoint-2010/ee556391(v=office.14)).


## Properties

- [Count](Access.Entities.Count.md)
- [Item](Access.Entities.Item.md)
- [Parent](Access.Entities.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]