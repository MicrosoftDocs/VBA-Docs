---
title: DoCmd.RepaintObject method (Access)
keywords: vbaac10.chm4169
f1_keywords:
- vbaac10.chm4169
ms.prod: access
api_name:
- Access.DoCmd.RepaintObject
ms.assetid: 6def040f-ae34-ce49-d3a0-786ad09bdc20
ms.date: 03/07/2019
localization_priority: Normal
---


# DoCmd.RepaintObject method (Access)

The **RepaintObject** method carries out the RepaintObject action in Visual Basic.


## Syntax

_expression_.**RepaintObject** (_ObjectType_, _ObjectName_)

_expression_ A variable that represents a **[DoCmd](Access.DoCmd.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**[AcObjectType](Access.AcObjectType.md)**|An **AcObjectType** constant that specifies the type of object to repaint.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _ObjectType_ argument.|

## Remarks

You can use the **RepaintObject** method to complete any pending screen updates for a specified database object or for the active database object, if none is specified. Such updates include any pending recalculations for the object's controls.

Using the **RepaintObject** method with no arguments (the default constant, **acDefault**, is assumed for the _ObjectType_ argument) repaints the active window.

The **RepaintObject** method of the **DoCmd** object was added to provide backwards compatibility for running the **RepaintObject** method in Visual Basic code in Microsoft Access 95. If you want to repaint a form, we recommend that you use the existing **Repaint** method of the **Form** object instead.


## Example

The following example repaints the **Customers** table. 

```vb
DoCmd.RepaintObject acTable, "Customers"
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]