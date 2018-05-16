---
title: AccessObject.Properties Property (Access)
keywords: vbaac10.chm12749
f1_keywords:
- vbaac10.chm12749
ms.prod: access
api_name:
- Access.AccessObject.Properties
ms.assetid: bfcf6d0a-3a1f-bd50-76c1-84a40b5dd769
ms.date: 06/08/2017
---


# AccessObject.Properties Property (Access)

Returns a reference to a  **[AccessObject](Access.AccessObject.md)** object's **[AccessObjectProperties](Access.AccessObjectProperties.md)** collection. Read-only.


## Syntax

 _expression_. **Properties**

 _expression_ A variable that represents an **AccessObject** object.


## Remarks

The  **AccessObjectProperties** collection object is the collection of all the properties related to a **AccessObject** object. You can refer to individual members of the collection by using the member object's index or a string expression that is the name of the member object. The first member object in the collection has an index value of 0 and the total number of member objects in the collection is the value of the **AccessObjectProperties** collection's **Count** property minus 1


## See also


#### Concepts


[AccessObject Object](Access.AccessObject.md)

