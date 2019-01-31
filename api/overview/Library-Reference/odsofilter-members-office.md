---
title: ODSOFilter members (Office)
ms.prod: office
ms.assetid: 2c4eeced-e51f-fbf9-65e5-93c06f099d58
ms.date: 01/30/2019
localization_priority: Normal
---


# ODSOFilter members (Office)

Represents a filter to be applied to an attached mail merge data source. The **ODSOFilter** object is a member of the **ODSOFilters** object.


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.ODSOFilter.Application.md)|Gets an **Application** object that represents the container application for the **ODSOFilter** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Column](../../Office.ODSOFilter.Column.md)|Gets or sets a **String** that represents the name of the field in the mail merge data source to use in the filter. Read/write.|
|[CompareTo](../../Office.ODSOFilter.CompareTo.md)|Gets or sets a **String** that represents the text to compare in the query filter criterion. Read/write.|
|[Comparison](../../Office.ODSOFilter.Comparison.md)|Gets or sets an **MsoFilterComparison** constant that represents how to compare the **Column** and **CompareTo** properties. Read/write.|
|[Conjunction](../../Office.ODSOFilter.Conjunction.md)|Gets or sets an **MsoFilterConjunction** constant that represents how a filter criterion relates to other filter criteria in the **ODSOFilters** object. Read/write.|
|[Creator](../../Office.ODSOFilter.Creator.md)|Gets a 32-bit integer that indicates the application in which the **ODSOFilter** object was created. Read-only.|
|[Index](../../Office.ODSOFilter.Index.md)|Gets a **Long** representing the index number for an **ODSOFilter** object in the collection. Read-only.|
|[Parent](../../Office.ODSOFilter.Parent.md)|Gets the **Parent** object for the **ODSOFilter** object. Read-only.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]