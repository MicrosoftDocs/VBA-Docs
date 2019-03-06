---
title: CustomControl object (Access)
keywords: vbaac10.chm12062
f1_keywords:
- vbaac10.chm12062
ms.prod: access
api_name:
- Access.CustomControl
ms.assetid: a6ded8cf-4cf8-26ff-bade-f37a7ac52b02
ms.date: 03/06/2019
localization_priority: Normal
---


# CustomControl object (Access)

When setting the properties of an ActiveX control, you may need or prefer to use the control's custom properties dialog box. This custom properties dialog box provides an alternative to the list of properties in the Microsoft Access property sheet for setting ActiveX control properties in Design view.


## Remarks

> [!NOTE] 
> This information only applies to ActiveX controls in a Microsoft Access database environment.

### Two ways to set properties

The reason for the custom properties dialog box is that not all applications that use ActiveX controls provide a property sheet like the one in Microsoft Access. The custom properties dialog box provides an interface for setting key control properties regardless of the interface supplied by the hosting application.

For some ActiveX control properties, you can choose either of these two locations to set the property:

- The Microsoft Access property sheet.  
- The ActiveX control's custom properties dialog box.
    
In some cases, the custom properties dialog box is the only way to set a property in Design view. This is usually the situation when the interface needed to set a property doesn't work inside the Microsoft Access property sheet. For example, the **GridFont** property for the Calendar control has a number of arguments; you can't set more than one argument per property in the Microsoft Access property sheet.

### Finding the custom properties dialog box

Not all ActiveX controls provide a custom properties dialog box. To see whether a control provides this custom properties dialog box, look for the **Custom** property in the Microsoft Access property sheet for this control. If the list of properties contains the name **Custom**, the control provides the custom properties dialog box.

After you choose the **Custom** property box in the Microsoft Access property sheet, choose the **Build** button to the right of the property box to display the control's custom properties dialog box, often presented as a tabbed dialog box. Choose the tab that contains the interface for setting the properties that you want to set.

### Using the custom properties dialog box

After you make changes on one tab, you can often apply those changes immediately by choosing the **Apply** button (if provided). You can choose other tabs to set other properties as needed. To approve all changes made in the custom properties dialog box, choose **OK**. To return to the Microsoft Access property sheet without changing any property settings, choose **Cancel**.

You can also view the custom properties dialog box by choosing the **Properties** subcommand of the ActiveX control **Object** command (for example, **Calendar Control Object**) on the **Edit** menu, or by choosing this same subcommand on the shortcut menu for the ActiveX control. In addition, some properties in the Microsoft Access property sheet for the ActiveX control, like the **GridFontColor** property of the Calendar control, have a **Build** button to the right of the property box. When you choose the **Build** button, the custom properties dialog box is displayed with the appropriate tab selected (for example, **Colors**).


## Events

- [Enter](Access.CustomControl.Enter.md)
- [Exit](Access.CustomControl.Exit.md)
- [GotFocus](Access.CustomControl.GotFocus.md)
- [LostFocus](Access.CustomControl.LostFocus.md)
- [Updated](Access.CustomControl.Updated.md)

## Methods

- [Move](Access.CustomControl.Move.md)
- [Requery](Access.CustomControl.Requery.md)
- [SetFocus](Access.CustomControl.SetFocus.md)
- [SizeToFit](Access.CustomControl.SizeToFit.md)

## Properties

- [About](Access.CustomControl.About.md)
- [Application](Access.CustomControl.Application.md)
- [BorderColor](Access.CustomControl.BorderColor.md)
- [BorderShade](Access.CustomControl.BorderShade.md)
- [BorderStyle](Access.CustomControl.BorderStyle.md)
- [BorderThemeColorIndex](Access.CustomControl.BorderThemeColorIndex.md)
- [BorderTint](Access.CustomControl.BorderTint.md)
- [BorderWidth](Access.CustomControl.BorderWidth.md)
- [BottomPadding](Access.CustomControl.BottomPadding.md)
- [Cancel](Access.CustomControl.Cancel.md)
- [Class](Access.CustomControl.Class.md)
- [Controls](Access.CustomControl.Controls.md)
- [ControlSource](Access.CustomControl.ControlSource.md)
- [ControlTipText](Access.CustomControl.ControlTipText.md)
- [ControlType](Access.CustomControl.ControlType.md)
- [Custom](Access.CustomControl.Custom.md)
- [Default](Access.CustomControl.Default.md)
- [DisplayWhen](Access.CustomControl.DisplayWhen.md)
- [Enabled](Access.CustomControl.Enabled.md)
- [EventProcPrefix](Access.CustomControl.EventProcPrefix.md)
- [GridlineColor](Access.CustomControl.GridlineColor.md)
- [GridlineStyleBottom](Access.CustomControl.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.CustomControl.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.CustomControl.GridlineStyleRight.md)
- [GridlineStyleTop](Access.CustomControl.GridlineStyleTop.md)
- [GridlineWidthBottom](Access.CustomControl.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.CustomControl.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.CustomControl.GridlineWidthRight.md)
- [GridlineWidthTop](Access.CustomControl.GridlineWidthTop.md)
- [Height](Access.CustomControl.Height.md)
- [HelpContextId](Access.CustomControl.HelpContextId.md)
- [HorizontalAnchor](Access.CustomControl.HorizontalAnchor.md)
- [InSelection](Access.CustomControl.InSelection.md)
- [IsVisible](Access.CustomControl.IsVisible.md)
- [Layout](Access.CustomControl.Layout.md)
- [LayoutID](Access.CustomControl.LayoutID.md)
- [Left](Access.CustomControl.Left.md)
- [LeftPadding](Access.CustomControl.LeftPadding.md)
- [Locked](Access.CustomControl.Locked.md)
- [Name](Access.CustomControl.Name.md)
- [Object](Access.CustomControl.Object.md)
- [ObjectPalette](Access.CustomControl.ObjectPalette.md)
- [ObjectVerbs](Access.CustomControl.ObjectVerbs.md)
- [ObjectVerbsCount](Access.CustomControl.ObjectVerbsCount.md)
- [OldBorderStyle](Access.CustomControl.OldBorderStyle.md)
- [OldValue](Access.CustomControl.OldValue.md)
- [OLEClass](Access.CustomControl.OLEClass.md)
- [OnEnter](Access.CustomControl.OnEnter.md)
- [OnExit](Access.CustomControl.OnExit.md)
- [OnGotFocus](Access.CustomControl.OnGotFocus.md)
- [OnLostFocus](Access.CustomControl.OnLostFocus.md)
- [OnUpdated](Access.CustomControl.OnUpdated.md)
- [Parent](Access.CustomControl.Parent.md)
- [Properties](Access.CustomControl.Properties.md)
- [RightPadding](Access.CustomControl.RightPadding.md)
- [Section](Access.CustomControl.Section.md)
- [SpecialEffect](Access.CustomControl.SpecialEffect.md)
- [TabIndex](Access.CustomControl.TabIndex.md)
- [TabStop](Access.CustomControl.TabStop.md)
- [Tag](Access.CustomControl.Tag.md)
- [Top](Access.CustomControl.Top.md)
- [TopPadding](Access.CustomControl.TopPadding.md)
- [Value](Access.CustomControl.Value.md)
- [VarOleObject](Access.CustomControl.VarOleObject.md)
- [Verb](Access.CustomControl.Verb.md)
- [VerticalAnchor](Access.CustomControl.VerticalAnchor.md)
- [Visible](Access.CustomControl.Visible.md)
- [Width](Access.CustomControl.Width.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)


[!include[Support and feedback](~/includes/feedback-boilerplate.md)]