---
title: Image object (Access)
keywords: vbaac10.chm10436
f1_keywords:
- vbaac10.chm10436
ms.prod: access
api_name:
- Access.Image
ms.assetid: 1bcc8552-94e2-b799-6903-392205cb4341
ms.date: 03/20/2019
localization_priority: Normal
---


# Image object (Access)

This object corresponds to an image control. The image control can add a picture to a form or report. For example, you could include an image control for a logo on an **Invoice** report.

> [!NOTE] 
> The functionality for the **Image** object's **Click** and **DoubleClick** events has been deprecated. If you want an image with click/double-click events, use instead a **Button** control and associate an image with that control to provide better accessibility. **Button** controls are part of the Tab Order loop, but **Image** controls are not. Existing applications will not be affected by this change.

## Remarks

|Control|Tool|
|:------|:----|
|![Image control](../images/t-imgctl_ZA06053959.gif)|![Image tool](../images/imagefrm_ZA06044465.gif)|

You can use the image control or an [Unbound object frame](overview/Access.md) for unbound pictures. The advantage of using the image control is that it's faster to display. The advantage of using the unbound object frame is that you can edit the object directly from the form or report.


## Events

- [Click](Access.Image.Click.md)
- [DblClick](Access.Image.DblClick.md)
- [MouseDown](Access.Image.MouseDown.md)
- [MouseMove](Access.Image.MouseMove.md)
- [MouseUp](Access.Image.MouseUp.md)

## Methods

- [Move](Access.Image.Move.md)
- [Requery](Access.Image.Requery.md)
- [SetFocus](Access.Image.SetFocus.md)
- [SizeToFit](Access.Image.SizeToFit.md)

## Properties

- [Application](Access.Image.Application.md)
- [BackColor](Access.Image.BackColor.md)
- [BackShade](Access.Image.BackShade.md)
- [BackStyle](Access.Image.BackStyle.md)
- [BackThemeColorIndex](Access.Image.BackThemeColorIndex.md)
- [BackTint](Access.Image.BackTint.md)
- [BorderColor](Access.Image.BorderColor.md)
- [BorderShade](Access.Image.BorderShade.md)
- [BorderStyle](Access.Image.BorderStyle.md)
- [BorderThemeColorIndex](Access.Image.BorderThemeColorIndex.md)
- [BorderTint](Access.Image.BorderTint.md)
- [BorderWidth](Access.Image.BorderWidth.md)
- [BottomPadding](Access.Image.BottomPadding.md)
- [Controls](Access.Image.Controls.md)
- [ControlTipText](Access.Image.ControlTipText.md)
- [ControlType](Access.Image.ControlType.md)
- [DisplayWhen](Access.Image.DisplayWhen.md)
- [EventProcPrefix](Access.Image.EventProcPrefix.md)
- [GridlineColor](Access.Image.GridlineColor.md)
- [GridlineShade](Access.Image.GridlineShade.md)
- [GridlineStyleBottom](Access.Image.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.Image.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.Image.GridlineStyleRight.md)
- [GridlineStyleTop](Access.Image.GridlineStyleTop.md)
- [GridlineThemeColorIndex](Access.Image.GridlineThemeColorIndex.md)
- [GridlineTint](Access.Image.GridlineTint.md)
- [GridlineWidthBottom](Access.Image.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.Image.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.Image.GridlineWidthRight.md)
- [GridlineWidthTop](Access.Image.GridlineWidthTop.md)
- [Height](Access.Image.Height.md)
- [HelpContextId](Access.Image.HelpContextId.md)
- [HorizontalAnchor](Access.Image.HorizontalAnchor.md)
- [Hyperlink](Access.Image.Hyperlink.md)
- [HyperlinkAddress](Access.Image.HyperlinkAddress.md)
- [HyperlinkSubAddress](Access.Image.HyperlinkSubAddress.md)
- [ImageHeight](Access.Image.ImageHeight.md)
- [ImageWidth](Access.Image.ImageWidth.md)
- [InSelection](Access.Image.InSelection.md)
- [IsVisible](Access.Image.IsVisible.md)
- [Layout](Access.Image.Layout.md)
- [LayoutID](Access.Image.LayoutID.md)
- [Left](Access.Image.Left.md)
- [LeftPadding](Access.Image.LeftPadding.md)
- [Name](Access.Image.Name.md)
- [ObjectPalette](Access.Image.ObjectPalette.md)
- [OldBorderStyle](Access.Image.OldBorderStyle.md)
- [OldValue](Access.Image.OldValue.md)
- [OnClick](Access.Image.OnClick.md)
- [OnDblClick](Access.Image.OnDblClick.md)
- [OnMouseDown](Access.Image.OnMouseDown.md)
- [OnMouseMove](Access.Image.OnMouseMove.md)
- [OnMouseUp](Access.Image.OnMouseUp.md)
- [Parent](Access.Image.Parent.md)
- [Picture](Access.Image.Picture.md)
- [PictureAlignment](Access.Image.PictureAlignment.md)
- [PictureData](Access.Image.PictureData.md)
- [PictureTiling](Access.Image.PictureTiling.md)
- [PictureType](Access.Image.PictureType.md)
- [Properties](Access.Image.Properties.md)
- [RightPadding](Access.Image.RightPadding.md)
- [Section](Access.Image.Section.md)
- [ShortcutMenuBar](Access.Image.ShortcutMenuBar.md)
- [SizeMode](Access.Image.SizeMode.md)
- [SpecialEffect](Access.Image.SpecialEffect.md)
- [Tag](Access.Image.Tag.md)
- [Top](Access.Image.Top.md)
- [TopPadding](Access.Image.TopPadding.md)
- [VerticalAnchor](Access.Image.VerticalAnchor.md)
- [Visible](Access.Image.Visible.md)
- [Width](Access.Image.Width.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
