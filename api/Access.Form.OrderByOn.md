---
title: Form.OrderByOn property (Access)
keywords: vbaac10.chm13349
f1_keywords:
- vbaac10.chm13349
ms.prod: access
api_name:
- Access.Form.OrderByOn
ms.assetid: 8902a8be-344e-d88f-8ac4-71d94dd0e3f0
ms.date: 03/14/2019
localization_priority: Normal
---


# Form.OrderByOn property (Access)

You can use the **OrderByOn** property to specify whether an object's **OrderBy** property setting is applied. Read/write **Boolean**.


## Syntax

_expression_.**OrderByOn**

_expression_ A variable that represents a **[Form](Access.Form.md)** object.


## Remarks

The **OrderByOn** property uses the following settings.

|Setting|Visual Basic|Description|
|:-----|:-----|:-----|
|Yes|**True**|The **OrderBy** property setting is applied when the object is opened.|
|No|**False**|(Default) The **OrderBy** property setting isn't applied when the object is opened.|

> [!NOTE] 
> When a new object is created, it inherits the **RecordSource**, **Filter**, **OrderBy**, **OrderByOn**, and **FilterOn** properties of the table or query that it was created from.


## Example

The following example displays a message indicating the state of the **OrderByOn** property for the **Mailing List** form.

```vb
MsgBox "OrderByOn property is " & Forms("Mailing List").OrderByOn
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]