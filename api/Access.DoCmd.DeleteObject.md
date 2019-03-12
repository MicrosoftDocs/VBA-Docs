---
title: DoCmd.DeleteObject method (Access)
keywords: vbaac10.chm4147
f1_keywords:
- vbaac10.chm4147
ms.prod: access
api_name:
- Access.DoCmd.DeleteObject
ms.assetid: 8e59c5a8-89bd-0d90-9fd1-a1178c73c1c1
ms.date: 03/06/2019
localization_priority: Normal
---


# DoCmd.DeleteObject method (Access)

The **DeleteObject** method carries out the DeleteObject action in Visual Basic.


## Syntax

_expression_.**DeleteObject** (_ObjectType_, _ObjectName_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that represents the type of object to delete.|
| _ObjectName_|Optional|**Variant**| A string expression that's the valid name of an object of the type selected by the _ObjectType_ argument. If you run Visual Basic code containing the **DeleteObject** method in a library database, Microsoft Access looks for the object with this name first in the library database, and then in the current database.|

## Remarks

You can use the **DeleteObject** method to delete a specified database object.

If you leave the _ObjectType_ and _ObjectName_ arguments blank (the default constant, **acDefault**, is assumed for _ObjectType_), Access deletes the object selected in the Database window. To select an object in the Database window, you can use the SelectObject action or **SelectObject** method with the _InDatabaseWindow_ argument set to Yes (**True**).


## Example

The following example deletes the specified table.

```vb
DoCmd.DeleteObject acTable, "Former Employees Table"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
