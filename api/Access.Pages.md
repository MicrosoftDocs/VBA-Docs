---
title: Pages object (Access)
keywords: vbaac10.chm10125
f1_keywords:
- vbaac10.chm10125
ms.prod: access
api_name:
- Access.Pages
ms.assetid: e77c8d31-1cb7-d647-6faa-2eb234ce0708
ms.date: 03/21/2019
localization_priority: Normal
---


# Pages object (Access)

The **Pages** collection contains all **[Page](Access.Page.md)** objects in a tab control.


## Remarks

The **Pages** collection is a special kind of **Controls** collection belonging to the tab control. It contains **Page** objects, which are controls. The **Pages** collection differs from a typical **Controls** collection in that you can add and remove **Page** objects by using methods of the **Pages** collection.

To add a new **Page** object to the **Pages** collection from Visual Basic, use the **Add** method. To remove an existing **Page** object, use the **Remove** method. To count the number of **Page** objects in the **Pages** collection, use the **Count** property.

You can also use the **[CreateControl](Access.Application.CreateControl.md)** method to add a **Page** object to the **Pages** collection of a tab control. To do this, you must specify the name of the tab control for the _Parent_ argument of the **CreateControl** function. The **[ControlType](Access.Page.ControlType.md)** property constant for a **Page** object is **acPage**.

You can enumerate through the **Pages** collection by using the **For Each...Next** statement.

Individual **Page** objects in the **Pages** collection are indexed beginning with zero.


## Methods

- [Add](Access.Pages.Add.md)
- [Remove](Access.Pages.Remove.md)

## Properties

- [Count](Access.Pages.Count.md)
- [Item](Access.pages.item.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]