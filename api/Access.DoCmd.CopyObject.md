---
title: DoCmd.CopyObject method (Access)
keywords: vbaac10.chm4146
f1_keywords:
- vbaac10.chm4146
ms.prod: access
api_name:
- Access.DoCmd.CopyObject
ms.assetid: 003e5b47-f8a2-2b6a-5e0c-7fb3e87b3258
ms.date: 03/06/2019
localization_priority: Normal
---


# DoCmd.CopyObject method (Access)

The **CopyObject** method carries out the CopyObject action in Visual Basic.


## Syntax

_expression_.**CopyObject** (_DestinationDatabase_, _NewName_, _SourceObjectType_, _SourceObjectName_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _DestinationDatabase_|Optional|**Variant**|A string expression that's the valid path and file name for the database that you want to copy the object into. To select the current database, leave this argument blank.<br/><br/>**NOTE**: In a Microsoft Access project (.adp), you must leave the _DestinationDatabase_ argument blank. If you execute Visual Basic code containing the **CopyObject** method in a library database and leave this argument blank, Access copies the object into the library database.|
| _NewName_|Optional|**Variant**|A string expression that's the new name for the object that you want to copy. To use the same name if you are copying into another database, leave this argument blank.|
| _SourceObjectType_|Optional|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that represents the type of object to copy.|
| _SourceObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the _SourceObjectType_ argument. If you run Visual Basic code containing the **CopyObject** method in a library database, Access looks for the object with this name first in the library database, and then in the current database.|

## Remarks

You can use the CopyObject action to copy the specified database object to a different Access database or to the same database or Access project (.adp) under a new name. For example, you can copy or back up an existing object in another database or quickly create a similar object with a few changes.

You must include either the _DestinationDatabase_ or _NewName_ argument or both for this method.

If you leave the _SourceObjectType_ and _SourceObjectName_ arguments blank (the default constant, **acDefault**, is assumed for _SourceObjectType_), Access copies the object selected in the Database window. To select an object in the Database window, you can use the SelectObject action or **SelectObject** method with the _InDatabaseWindow_ argument set to Yes (**True**).

If you specify the _SourceObjectType_ and _SourceObjectName_ arguments but leave either the _NewName_ argument or the _DestinationDatabase_ argument blank, you must include the _NewName_ or _DestinationDatabase_ argument's comma. If you leave a trailing argument blank, don't use a comma following the last argument that you specify.


## Example

The following example uses the **CopyObject** method to copy the **Employees** table and give it a new name in the current database.

```vb
DoCmd.CopyObject, "Employees Copy", acTable, "Employees"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
