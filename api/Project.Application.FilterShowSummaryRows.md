---
title: Application.FilterShowSummaryRows method (Project)
keywords: vbapj.chm2297
f1_keywords:
- vbapj.chm2297
ms.prod: project-server
api_name:
- Project.Application.FilterShowSummaryRows
ms.assetid: 173bf591-7579-505f-3cbd-42eaddb231ad
ms.date: 06/08/2017
localization_priority: Normal
---


# Application.FilterShowSummaryRows method (Project)

Shows or hides the related summary rows.


## Syntax

_expression_. `FilterShowSummaryRows`( `_On_` )

 _expression_ An expression that returns an **[Application](Project.Application.md)** object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _On_|Required|**Boolean**|**True** if summary rows are shown; otherwise, **False**.|

## Return value

 **Boolean**


## Remarks

The **FilterShowSummaryRows** method corresponds to the following command on the ribbon: on the **View** tab, click the **Filter** drop-down list box in the **Data** section, and then click **Show Related Summary Rows**.


## Example

If the current filter shows only completed tasks, for example, the following command shows the summary tasks.


```vb
FilterShowSummaryRows (true)
```

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]