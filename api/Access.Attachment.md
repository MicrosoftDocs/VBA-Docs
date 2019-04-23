---
title: Attachment object (Access)
keywords: vbaac10.chm14036
f1_keywords:
- vbaac10.chm14036
ms.prod: access
api_name:
- Access.Attachment
ms.assetid: b0756145-9012-f9b9-7df9-e168defed3bf
ms.date: 02/07/2019
localization_priority: Normal
---


# Attachment object (Access)

This object corresponds to an attachment control. Use an attachment control when you want to manipulate the contents fields of the attachment data type.


## Remarks

> [!NOTE] 
> You can attach files only to databases that you create in Office Access 2007 and later and that use the new .accdb file format. You cannot share attachments between an Office Access 2007 (.accdb) database and a database in the earlier (.mdb) file format.

You can attach a maximum of two gigabytes of data (the maximum size for an Access database). Individual files cannot exceed 256 megabytes in size.


### Supported image file formats

Office Access 2007 and later support the following graphic file formats natively, meaning the attachment control renders them without the need for additional software.

- BMP (Windows Bitmap)   
- RLE (Run Length Encoded Bitmap)   
- DIB (Device Independent Bitmap)    
- GIF (Graphics Interchange Format)    
- JPEG, JPG, JPE (Joint Photographic Experts Group)    
- EXIF (Exchangeable File Format)    
- PNG (Portable Network Graphics)    
- TIFF, TIF (Tagged Image File Format)    
- ICON, ICO (Icon)    
- WMF (Windows Metafile)    
- EMF (Enhanced Metafile)
    

### Supported formats for documents and other files

As a rule, you can attach any file that was created with one of the 2007 Microsoft Office or later system programs. You can also attach log files (.log), text files (.text, .txt), and compressed .zip files.


### File-naming conventions

The names of your attached files can contain any Unicode character supported by the NTFS file system used in Microsoft Windows NT (NTFS). In addition, file names must conform to these guidelines:

- Names must not exceed 255 characters, including the file name extensions.
    
- Names cannot contain the following characters: question marks (?), quotation marks ("), forward or backward slashes (/ \\), opening or closing brackets (< >), asterisks (*), vertical bars or pipes ( | ), colons ( : ), or paragraph marks.
    

### Types of files that Access compresses

Access will compress your attached files unless those files are compressed natively. For example, JPEG files are compressed by the graphics program that created them, so Access does not compress them. The following table lists some supported file types and whether or not Access compresses them.

|File extension|Compressed?|Reason|
|:-----|:-----|:-----|
|.jpg, .jpeg|No|Already compressed|
|.gif|No|Already compressed|
|.png|No|Already compressed|
|.tif, .tiff|Yes||
|.exif|Yes||
| .bmp|Yes||
|.emf|Yes||
|.wmf|Yes||
|.ico|Yes||
|.zip|No|Already compressed|
|.cab|No|Already compressed|
|.docx|No|Already compressed|
|.xlsx|No|Already compressed|
|.xlsb|No|Already compressed|
|.pptx|No|Already compressed|

### Blocked file formats

Office Access 2007 blocks the following types of attached files. At this time, you cannot unblock any of the file types listed here.

|||||
|:-----|:-----|:-----|:-----|
|.ade|.ins|.mda|.scr|
|.adp|.isp|.mdb|.sct|
|.app|.its|.mde|.shb|
|.asp|.js |.mdt|.shs|
|.bas|.jse|.mdw|.tmp|
|.bat|.ksh|.mdz|.url|
|.cer|.lnk|.msc|.vb|
|.chm|.mad|.msi|.vbe|
|.cmd|.maf|.msp|.vbs|
|.com|.mag|.mst|.vsmacros|
|.cpl|.mam|.ops|.vss|
|.crt|.maq|.pcd|.vst|
|.csh|.mar|.pif|.vsw|
|.exe|.mas|.prf|.ws|
|.fxp|.mat|.prg|.wsc|
|.hlp|.mau|.pst|.wsf|
|.hta|.mav|.reg|.wsh|
|.inf|.maw|.scf||

## Events

- [AfterUpdate](Access.Attachment.AfterUpdate-event.md)
- [AttachmentCurrent](Access.Attachment.AttachmentCurrent.md)
- [BeforeUpdate](Access.Attachment.BeforeUpdate-event.md)
- [Change](Access.Attachment.Change.md)
- [Click](Access.Attachment.Click.md)
- [DblClick](Access.Attachment.DblClick.md)
- [Dirty](Access.Attachment.Dirty.md)
- [Enter](Access.Attachment.Enter.md)
- [Exit](Access.Attachment.Exit.md)
- [GotFocus](Access.Attachment.GotFocus.md)
- [KeyDown](Access.Attachment.KeyDown.md)
- [KeyPress](Access.Attachment.KeyPress.md)
- [KeyUp](Access.Attachment.KeyUp.md)
- [LostFocus](Access.Attachment.LostFocus.md)
- [MouseDown](Access.Attachment.MouseDown.md)
- [MouseMove](Access.Attachment.MouseMove.md)
- [MouseUp](Access.Attachment.MouseUp.md)

## Methods

- [Back](Access.Attachment.Back.md)
- [Forward](Access.Attachment.Forward.md)
- [Move](Access.Attachment.Move.md)
- [Requery](Access.Attachment.Requery.md)
- [SetFocus](Access.Attachment.SetFocus.md)
- [SizeToFit](Access.Attachment.SizeToFit.md)

## Properties

- [AddColon](Access.Attachment.AddColon.md)
- [AfterUpdate](Access.Attachment.AfterUpdate-property.md)
- [Application](Access.Attachment.Application.md)
- [AttachmentCount](Access.Attachment.AttachmentCount.md)
- [AutoLabel](Access.Attachment.AutoLabel.md)
- [BackColor](Access.Attachment.BackColor.md)
- [BackShade](Access.Attachment.BackShade.md)
- [BackStyle](Access.Attachment.BackStyle.md)
- [BackThemeColorIndex](Access.Attachment.BackThemeColorIndex.md)
- [BackTint](Access.Attachment.BackTint.md)
- [BeforeUpdate](Access.Attachment.BeforeUpdate-property.md)
- [BorderColor](Access.Attachment.BorderColor.md)
- [BorderShade](Access.Attachment.BorderShade.md)
- [BorderStyle](Access.Attachment.BorderStyle.md)
- [BorderThemeColorIndex](Access.Attachment.BorderThemeColorIndex.md)
- [BorderTint](Access.Attachment.BorderTint.md)
- [BorderWidth](Access.Attachment.BorderWidth.md)
- [BottomPadding](Access.Attachment.BottomPadding.md)
- [ColumnHidden](Access.Attachment.ColumnHidden.md)
- [ColumnOrder](Access.Attachment.ColumnOrder.md)
- [ColumnWidth](Access.Attachment.ColumnWidth.md)
- [Controls](Access.Attachment.Controls.md)
- [ControlSource](Access.Attachment.ControlSource.md)
- [ControlTipText](Access.Attachment.ControlTipText.md)
- [ControlType](Access.Attachment.ControlType.md)
- [CurrentAttachment](Access.Attachment.CurrentAttachment.md)
- [DefaultPicture](Access.Attachment.DefaultPicture.md)
- [DefaultPictureType](Access.Attachment.DefaultPictureType.md)
- [DisplayAs](Access.Attachment.DisplayAs.md)
- [DisplayWhen](Access.Attachment.DisplayWhen.md)
- [Enabled](Access.Attachment.Enabled.md)
- [EventProcPrefix](Access.Attachment.EventProcPrefix.md)
- [FileName](Access.Attachment.FileName.md)
- [FileType](Access.Attachment.FileType.md)
- [FileURL](Access.Attachment.FileURL.md)
- [GridlineColor](Access.Attachment.GridlineColor.md)
- [GridlineShade](Access.Attachment.GridlineShade.md)
- [GridlineStyleBottom](Access.Attachment.GridlineStyleBottom.md)
- [GridlineStyleLeft](Access.Attachment.GridlineStyleLeft.md)
- [GridlineStyleRight](Access.Attachment.GridlineStyleRight.md)
- [GridlineStyleTop](Access.Attachment.GridlineStyleTop.md)
- [GridlineThemeColorIndex](Access.Attachment.GridlineThemeColorIndex.md)
- [GridlineTint](Access.Attachment.GridlineTint.md)
- [GridlineWidthBottom](Access.Attachment.GridlineWidthBottom.md)
- [GridlineWidthLeft](Access.Attachment.GridlineWidthLeft.md)
- [GridlineWidthRight](Access.Attachment.GridlineWidthRight.md)
- [GridlineWidthTop](Access.Attachment.GridlineWidthTop.md)
- [Height](Access.Attachment.Height.md)
- [HelpContextId](Access.Attachment.HelpContextId.md)
- [HorizontalAnchor](Access.Attachment.HorizontalAnchor.md)
- [InSelection](Access.Attachment.InSelection.md)
- [IsVisible](Access.Attachment.IsVisible.md)
- [LabelAlign](Access.Attachment.LabelAlign.md)
- [LabelX](Access.Attachment.LabelX.md)
- [LabelY](Access.Attachment.LabelY.md)
- [Layout](Access.Attachment.Layout.md)
- [LayoutID](Access.Attachment.LayoutID.md)
- [Left](Access.Attachment.Left.md)
- [LeftPadding](Access.Attachment.LeftPadding.md)
- [Locked](Access.Attachment.Locked.md)
- [Name](Access.Attachment.Name.md)
- [OldBorderStyle](Access.Attachment.OldBorderStyle.md)
- [OldValue](Access.Attachment.OldValue.md)
- [OnAttachmentCurrent](Access.Attachment.OnAttachmentCurrent.md)
- [OnChange](Access.Attachment.OnChange.md)
- [OnClick](Access.Attachment.OnClick.md)
- [OnDblClick](Access.Attachment.OnDblClick.md)
- [OnDirty](Access.Attachment.OnDirty.md)
- [OnEnter](Access.Attachment.OnEnter.md)
- [OnExit](Access.Attachment.OnExit.md)
- [OnGotFocus](Access.Attachment.OnGotFocus.md)
- [OnKeyDown](Access.Attachment.OnKeyDown.md)
- [OnKeyPress](Access.Attachment.OnKeyPress.md)
- [OnKeyUp](Access.Attachment.OnKeyUp.md)
- [OnLostFocus](Access.Attachment.OnLostFocus.md)
- [OnMouseDown](Access.Attachment.OnMouseDown.md)
- [OnMouseMove](Access.Attachment.OnMouseMove.md)
- [OnMouseUp](Access.Attachment.OnMouseUp.md)
- [Parent](Access.Attachment.Parent.md)
- [PictureAlignment](Access.Attachment.PictureAlignment.md)
- [PictureSizeMode](Access.Attachment.PictureSizeMode.md)
- [PictureTiling](Access.Attachment.PictureTiling.md)
- [Properties](Access.Attachment.Properties.md)
- [RightPadding](Access.Attachment.RightPadding.md)
- [Section](Access.Attachment.Section.md)
- [ShortcutMenuBar](Access.Attachment.ShortcutMenuBar.md)
- [SpecialEffect](Access.Attachment.SpecialEffect.md)
- [StatusBarText](Access.Attachment.StatusBarText.md)
- [TabIndex](Access.Attachment.TabIndex.md)
- [TabStop](Access.Attachment.TabStop.md)
- [Tag](Access.Attachment.Tag.md)
- [Top](Access.Attachment.Top.md)
- [TopPadding](Access.Attachment.TopPadding.md)
- [VerticalAnchor](Access.Attachment.VerticalAnchor.md)
- [Visible](Access.Attachment.Visible.md)
- [Width](Access.Attachment.Width.md)

## See also

- [Access Object Model Reference](overview/Access/object-model.md)



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
