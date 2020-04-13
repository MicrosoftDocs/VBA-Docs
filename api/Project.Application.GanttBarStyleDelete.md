---
title: Application.GanttBarStyleDelete method (Project)
keywords: vbapj.chm2059
f1_keywords:
- vbapj.chm2059
ms.prod: project-server
api_name:
- Project.Application.GanttBarStyleDelete
ms.assetid: 3cac2b37-147c-f1bf-bc94-d2bc9bffa14b
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.GanttBarStyleDelete method (Project)

Deletes a Gantt bar style from the active Gantt Chart.


## Syntax

_expression_. `GanttBarStyleDelete`( `_Item_` )

_expression_ A variable that represents an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Item_|Required|**String**|**String**. The name or row number of the Gantt bar to delete from the **Bar Styles** dialog box.|

## Return value

 **Boolean**


## Remarks

To manually show the **Bar Styles** dialog box, click the **Format** tab under the **Gantt Chart Tools** tab. In the **Bar Styles** group, click **Bar Styles** in the **Format** drop-down list. The **Bar Styles** dialog box can contain up to 200 style entries.


## Example

The following command deletes style number 41 in the **Bar Styles** dialog box.


```vb
GanttBarStyleDelete Item:="41"
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]