---
title: Shapes.Item method (Project)
ms.prod: project-server
ms.assetid: 43fba4f4-f3d3-20a0-2c77-15e31dcdcbf5
ms.date: 06/08/2017
localization_priority: Normal
---


# Shapes.Item method (Project)
Returns an individual  **Shape** object in the **Shapes** collection.

## Syntax

_expression_.**Item** (_Index_)

_expression_ A variable that represents a **[Shapes](Project.Shapes.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Index_|Required|**Variant**|Can be a  **String** value for the name of the shape or a **Long** value for the ordinal index number of the shape.|
| _Index_|Required|**Variant**||

## Return value

 **Shape**

The shape that is specified by the  _Index_ parameter.


## Remarks

The  **Item** method acts like the default **[Shapes.Value](Project.shapes.value.md)** property. For example, create a report namedTable Tests that contains a table. The following statement in the Immediate window of the VBE prints the name of the table.


```vb
? ActiveProject.Reports("Table Tests").Shapes.Item(1).Name
```

If you leave off the  **Item** method, the following statement has the same output, but uses the default **Value** property to get the **Shape** object.




```vb
? ActiveProject.Reports("Table Tests").Shapes(1).Name
```

The following statement is the same as the previous:




```vb
? ActiveProject.Reports("Table Tests").Shapes.Value(1).Name
```


## See also


[Shapes Object](Project.shapes.md)
[Value Property](Project.shapes.value.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]