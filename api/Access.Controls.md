---
title: Controls object (Access)
keywords: vbaac10.chm10176
f1_keywords:
- vbaac10.chm10176
ms.prod: access
api_name:
- Access.Controls
ms.assetid: 26771888-86e8-28c3-6668-f793474cbb5b
ms.date: 03/06/2019
localization_priority: Normal
---


# Controls object (Access)

The **Controls** collection contains all of the controls on a form, report, or subform, within another control, or attached to another control. The **Controls** collection is a member of the **[Form](Access.Form.md)**, **[Report](Access.Report.md)**, and **[SubForm](Access.SubForm.md)** objects.


## Remarks

You can enumerate individual controls, count them, and set their properties in the **Controls** collection. For example, you can enumerate the **Controls** collection of a particular form and set the **Height** property of each control to a specified value.

It is faster to refer to the **Controls** collection implicitly, as in the following examples, which refer to a control called **NewData** on a form named **OrderForm**. Of the following syntax examples, `Me!NewData` is the fastest way to refer to the control.

```vb
Me!NewData               ' Or Forms!OrderForm!NewData.
```


```vb
Me![New Data]            ' Use if control name contains space.
```


```vb
Me("NewData")            ' Performance is slightly slower.
```

<br/>

You can also refer to an individual control by referring explicitly to the **Controls** collection.

```vb
Me.Controls!NewData      ' Or Forms!OrderForm.Controls!NewData.
```


```vb
Me.Controls![New Data]
```


```vb
Me.Controls("NewData")
```

<br/>

Additionally, you can refer to a control by its index in the collection. The **Controls** collection is indexed beginning with zero.

```vb
Me(0)                    ' Refer to first item in collection.
```


```vb
Me.Controls(0)
```

> [!NOTE] 
> You can use the **Me** keyword to represent a form or report within code only if you are referring to the form or report from code within the form module or report module. If you are referring to a form or report from a standard module or a different form's or report's module, you must use the full reference to the form or report.

To work with the controls on a section of a form or report, use the **Section** property to return a reference to a **Section** object. You can then refer to the **Controls** collection of the **Section** object.

Two types of **Control** objects, the tab control and option group control, have **Controls** collections that can contain multiple controls. The **Controls** collection belonging to the option group control contains any option button, check box, toggle button, or label controls in the option group.

The tab control contains a **[Pages](Access.Pages.md)** collection, which is a special type of **Controls** collection. The **Pages** collection contains **[Page](Access.Page.md)** objects. **Page** objects are also controls. The **[ControlType](Access.Page.ControlType.md)** property constant for a **Page** control is **acPage**. A **Page** object, in turn, has its own **Controls** collection, which contains all the controls on an individual page.

Other **Control** objects have a **Controls** collection that can contain an attached label. These controls include the text box, option group, option button, toggle button, check box, combo box, list box, command button, bound object frame, and unbound object frame controls.


## Properties

- [Application](Access.Controls.Application.md)
- [Count](Access.Controls.Count.md)
- [Item](Access.Controls.Item.md)
- [Parent](Access.Controls.Parent.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
