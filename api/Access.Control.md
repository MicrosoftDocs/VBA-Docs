---
title: Control object (Access)
keywords: vbaac10.chm10174
f1_keywords:
- vbaac10.chm10174
ms.prod: access
api_name:
- Access.Control
ms.assetid: ce2362e5-4390-590e-06c0-6f27e8d988cd
ms.date: 03/06/2019
localization_priority: Normal
---


# Control object (Access)

The **Control** object represents a control on a form, report, or section, within another control, or attached to another control.


## Remarks

All controls on a form or report belong to the **Controls** collection for that **Form** or **Report** object. Controls within a particular section belong to the **Controls** collection for that section. Controls within a tab control or option group control belong to the **Controls** collection for that control. A label control that is attached to another control belongs to the **Controls** collection for that control.

When you refer to an individual **Control** object in the **Controls** collection, you can refer to the **Controls** collection either implicitly or explicitly.

```vb
' Implicitly refer to NewData control in Controls 
' collection. 
Me!NewData
```


```vb
' Use if control name contains space. 
Me![New Data]
```


```vb
' Performance slightly slower. 
Me("NewData")
```


```vb
' Refer to a control by its index in the controls 
' collection. 
Me(0)
```


```vb
' Refer to a NewData control by using the subform 
' Controls collection. 
Me.ctlSubForm.Controls!NewData
```


```vb
' Explicitly refer to the NewData control in the 
' Controls collection. 
Me.Controls!NewData
```


```vb
Me.Controls("NewData")
```


```vb
Me.Controls(0)
```

> [!NOTE] 
> You can use the **Me** keyword to represent a **Form** or **Report** object within code only if you are referring to the form or report from code within the class module. If you are referring to a form or report from a standard module or a different form's or report's module, you must use the full reference to the form or report.

Each **Control** object is denoted by a particular intrinsic constant. For example, the intrinsic constant **acTextBox** is associated with a text box control, and **acCommandButton** is associated with a command button. The constants for the various Microsoft Access controls are set forth in the control's **ControlType** property.

To determine the type of an existing control, you can use the **ControlType** property. However, you don't need to know the specific type of control to use it in code. You can simply represent it with a variable of data type **Control**.

If you do know the data type of the control to which you are referring, and the control is a built-in Microsoft Access control, you should represent it with a variable of a specific type. For example, if you know that a particular control is a text box, declare a variable of type **TextBox** to represent it, as shown in the following code.

```vb
Dim txt As TextBox 
Set txt = Forms!Employees!LastName 

```

> [!NOTE] 
> If a control is an ActiveX control, you must declare a variable of type **Control** to represent it; you cannot use a specific type. If you are not certain what type of control a variable will point to, declare the variable as type **Control**.

The option group control can contain other controls within its **Controls** collection, including option button, check box, toggle button, and label controls.

The tab control contains a **[Pages](Access.Pages.md)** collection, which is a special type of **Controls** collection. The **Pages** collection contains **[Page](Access.Page.md)** objects, which are controls. Each **Page** object in turn contains a **Controls** collection, which contains all of the controls on that page.

Other **Control** objects have a **Controls** collection that can contain an attached label. These controls include the text box, option group, option button, toggle button, check box, combo box, list box, command button, bound object frame, and unbound object frame controls.


## Methods

- [Dropdown](Access.Control.Dropdown.md)
- [Move](Access.Control.Move.md)
- [Requery](Access.Control.Requery.md)
- [SetFocus](Access.Control.SetFocus.md)
- [SizeToFit](Access.Control.SizeToFit.md)
- [Undo](Access.Control.Undo.md)

## Properties

- [Application](Access.Control.Application.md)
- [BottomPadding](Access.Control.BottomPadding.md)
- [Column](Access.Control.Column.md)
- [Controls](Access.Control.Controls.md)
- [Form](Access.Control.Form.md)
- [GridlineColor](Access.Control.GridlineColor.md)
- [GridlineStyleBottom](Access.Control.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.Control.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.Control.GridlineStyleRight.md)
- [GridlineStyleTop](Access.Control.GridlineStyleTop.md)
- [GridlineWidthBottom](Access.Control.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.Control.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.Control.GridlineWidthRight.md)
- [GridlineWidthTop](Access.Control.GridlineWidthTop.md)
- [HorizontalAnchor](Access.Control.HorizontalAnchor.md)
- [Hyperlink](Access.Control.Hyperlink.md)
- [ItemData](Access.Control.ItemData.md)
- [ItemsSelected](Access.Control.ItemsSelected.md)
- [Layout](Access.Control.Layout.md)
- [LayoutID](Access.Control.LayoutID.md)
- [LeftPadding](Access.Control.LeftPadding.md)
- [Name](Access.Control.Name.md)
- [Object](Access.Control.Object.md)
- [ObjectVerbs](Access.Control.ObjectVerbs.md)
- [OldValue](Access.Control.OldValue.md)
- [Pages](Access.Control.Pages.md)
- [Parent](Access.Control.Parent.md)
- [Properties](Access.Control.Properties.md)
- [Report](Access.Control.Report.md)
- [RightPadding](Access.Control.RightPadding.md)
- [Selected](Access.Control.Selected.md)
- [SmartTags](Access.Control.SmartTags.md)
- [TopPadding](Access.Control.TopPadding.md)
- [VerticalAnchor](Access.Control.VerticalAnchor.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
