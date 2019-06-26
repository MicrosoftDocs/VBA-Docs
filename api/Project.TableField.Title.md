---
title: TableField.Title property (Project)
keywords: vbapj.chm132689
f1_keywords:
- vbapj.chm132689
ms.prod: project-server
api_name:
- Project.TableField.Title
ms.assetid: 19ee2239-0a1c-73ca-9ea4-21fdfc924d65
ms.date: 06/08/2017
localization_priority: Normal
---


# TableField.Title property (Project)

Gets or sets the title of the field in a table. Read/write  **String**.


## Syntax

_expression_.**Title**

_expression_ A variable that represents a [TableField](./Project.TableField.md) object.


## Remarks

 **Title** is the default property of the **TableField** object.


> [!NOTE] 
> Many of the fields in a table do not have a default title, so the  **Title** property is an empty string ("").


## Example

The following statement prints "Task Name" in the  **Immediate** pane.


```vb
Debug.Print ActiveProject.TaskTables("Entry").TableFields(4)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]