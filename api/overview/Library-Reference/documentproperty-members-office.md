---
title: DocumentProperty members (Office)
ms.prod: office
ms.assetid: 568da0ff-fa90-150a-06ec-611de886334e
ms.date: 01/30/2019
localization_priority: Normal
---


# DocumentProperty members (Office)

Represents a custom or built-in document property of a container document. The **DocumentProperty** object is a member of the **DocumentProperties** collection.


## Methods

|Name|Description|
|:-----|:-----|
|[Delete](../../Office.DocumentProperty.Delete.md)|Removes a custom document property.|


## Properties

|Name|Description|
|:-----|:-----|
|[Application](../../Office.DocumentProperty.Application.md)|Gets an **Application** object that represents the container application for the **DocumentProperty** object (you can use this property with an **Automation** object to return that object's container application). Read-only.|
|[Creator](../../Office.DocumentProperty.Creator.md)|Gets a 32-bit integer that indicates the application in which the **DocumentProperty** object was created. Read-only.|
|[LinkSource](../../Office.DocumentProperty.LinkSource.md)|Gets or sets the source of a linked custom document property. Read/write.|
|[LinkToContent](../../Office.DocumentProperty.LinkToContent.md)|Is **True** if the value of the custom document property is linked to the content of the container document. **False** if the value is static. Read/write.|
|[Name](../../Office.DocumentProperty.Name.md)|Gets or sets the name of a document property. Read/write.|
|[Parent](../../Office.DocumentProperty.Parent.md)|Gets the **Parent** object for the **DocumentProperty** object. Read-only.|
|[Type](../../Office.DocumentProperty.Type.md)|Gets or sets the document property type. Read-only for built-in document properties; read/write for custom document properties.|
|[Value](../../Office.DocumentProperty.Value.md)|Gets or sets the value of a document property. Read/write.|

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]