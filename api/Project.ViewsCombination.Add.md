---
title: ViewsCombination.Add method (Project)
keywords: vbapj.chm132807
f1_keywords:
- vbapj.chm132807
ms.prod: project-server
api_name:
- Project.ViewsCombination.Add
ms.assetid: 84e93698-88c3-b4a7-a754-8078fcab897a
ms.date: 06/08/2017
localization_priority: Normal
---


# ViewsCombination.Add method (Project)

Adds a **ViewCombination** object to a **ViewsCombination** collection.


## Syntax

_expression_.**Add** (_Name_, _TopView_, _BottomView_, _ShowInMenu_)

_expression_ A variable that represents a 'ViewsCombination' object.


## Parameters



|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _Name_|Required|**String**|The name of the combination view.|
| _TopView_|Required|**Variant**|The view that appears in the top pane of a combination view.|
| _BottomView_|Required|**Variant**|The view that appears in the bottom pane of a combination view.|
| _ShowInMenu_|Optional|**Boolean**|**True** if Project Server shows the view in the **View** menu. The default value is **False**|

## Return value

 **ViewCombination**


## See also


[ViewsCombination Collection Object](Project.viewscombination(object).md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]