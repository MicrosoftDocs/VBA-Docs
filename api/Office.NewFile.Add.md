---
title: NewFile.Add method (Office)
keywords: vbaof11.chm235001
f1_keywords:
- vbaof11.chm235001
ms.prod: office
api_name:
- Office.NewFile.Add
ms.assetid: 094e4093-fc2d-beaa-4a63-b3ad88557907
ms.date: 01/22/2019
localization_priority: Normal
---


# NewFile.Add method (Office)

Adds a new item to the **New Item** task pane. Returns a **Boolean** value to indicate whether the operation was successful.


## Syntax

_expression_.**Add** (_FileName_, _Section_, _DisplayName_, _Action_)

_expression_ Required. A variable that represents a **[NewFile](Office.NewFile.md)** object.


## Parameters

|Name|Required/Optional|Data type|Description|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to add to the list of files on the task pane.|
| _Section_|Optional|**Variant**|The section to which to add the file. Can be any **[msoFileNew](office.msofilenewsection.md)** constant.|
| _DisplayName_|Optional|**Variant**|The text to display in the task pane.|
| _Action_|Optional|**Variant**|The action to take when a user clicks the item. Can be any **msoFileNew** constant.|

## See also

- [NewFile object members](overview/library-reference/newfile-members-office.md)




[!include[Support and feedback](~/includes/feedback-boilerplate.md)]