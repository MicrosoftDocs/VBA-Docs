---
title: Application.ProjectMove Method (Project)
keywords: vbapj.chm2291
f1_keywords:
- vbapj.chm2291
ms.prod: project-server
api_name:
- Project.Application.ProjectMove
ms.assetid: ba30bd12-a26a-12e5-8cff-df1a34a58df0
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.ProjectMove Method (Project)

Moves the project start date to a new date.


## Syntax

 _expression_. `ProjectMove`( ` _Date_`, ` _MoveDeadline_` )

 _expression_ An expression that returns an [Application](./Project.Application.md) object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Date_|Optional|**Variant**|Specifies the new project start date.|
| _MoveDeadline_|Optional|**Boolean**|**True** if deadlines are also moved; otherwise **false**. The default is **true**.|

## Return value

 **Boolean**


## Remarks

The  **ProjectMove** method is equivalent to clicking **Move Project** on the **Project** tab of the Ribbon.

Running the  **ProjectMove** method with no arguments displays the **Move Project** dialog box.


## Example

The following command moves the project start date to May 23, 2012. Deadlines are moved a corresponding number of days.


```vb
projectmove Date:="5/23/2012"
```


