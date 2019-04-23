---
title: Application.References property (Access)
keywords: vbaac10.chm12564
f1_keywords:
- vbaac10.chm12564
ms.prod: access
api_name:
- Access.Application.References
ms.assetid: da78f26f-1127-796d-bba1-f1c0d98a582e
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.References property (Access)

You can use the **References** property to access the **[References](Access.References.md)** collection and its related properties, methods, and events. Read-only **References** collection.


## Syntax

_expression_.**References**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

The **References** collection corresponds to the list of references in the **References** dialog box, available by clicking **References** on the **Tools** menu. Each **Reference** object represents one selected reference in the list. References that appear in the **References** dialog box but haven't been selected aren't in the **References** collection.
    

## Example

The following example displays a message indicating the number of boxes selected in the **References** dialog box.

```vb
MsgBox "There are " & Application.References.Count & " references."
```



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]