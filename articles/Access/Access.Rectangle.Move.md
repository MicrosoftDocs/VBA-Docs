---
title: Rectangle.Move Method (Access)
keywords: vbaac10.chm10306
f1_keywords:
- vbaac10.chm10306
ms.prod: access
api_name:
- Access.Rectangle.Move
ms.assetid: d5550716-69f9-fcdd-3547-1178b2c1a9f0
ms.date: 06/08/2017
---


# Rectangle.Move Method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

 _expression_. **Move**( ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents a **Rectangle** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Variant**|The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.|
| _Top_|Optional|**Variant**|The screen position in twips for the top edge of the object relative to the top edge of the Microsoft Access window.|
| _Width_|Optional|**Variant**|The desired width in twips of the object.|
| _Height_|Optional|**Variant**|The desired height in twips of the object.|

## Remarks

Only the  _Left_ argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specify _Width_ without specifying _Left_ and _Top_. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## See also


#### Concepts


[Rectangle Object](Access.Rectangle.md)

