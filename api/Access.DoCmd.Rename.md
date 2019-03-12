---
title: DoCmd.Rename method (Access)
keywords: vbaac10.chm4168
f1_keywords:
- vbaac10.chm4168
ms.prod: access
api_name:
- Access.DoCmd.Rename
ms.assetid: c9286727-a172-b7c5-c8b4-6e63012db98a
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.Rename method (Access)

The **Rename** method carries out the Rename action in Visual Basic.


## Syntax

_expression_.**Rename** (_NewName_, _ObjectType_, _OldName_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _NewName_|Required|**Variant**| A string expression that's the new name for the object that you want to rename. The name must follow the object-naming rules for Microsoft Access objects.|
| _ObjectType_|Optional|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that specifies the type of object to rename. The default value is **acDefault**.|
| _OldName_|Optional|**Variant**| A string expression that's the valid name of an object of the type specified by the _ObjectType_ argument. If you execute Visual Basic code containing the **Rename** method in a library database, Access looks for the object with this name first in the library database, and then in the current database.|

## Remarks

You can use the **Rename** method to rename a specified database object.

If you leave the _ObjectType_ and _OldName_ arguments blank (the default constant, **acDefault**, is assumed for _ObjectType_), Access renames the object selected in the Database window. To select an object in the Database window, you can use the **SelectObject** method with the _InDatabaseWindow_ argument set to Yes (**True**).


## Example

The following example renames the **Employees** table.

```vb
DoCmd.Rename "Old Employees Table", acTable, "Employees"
```




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
