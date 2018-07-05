---
title: Page Members (Outlook Forms Script)
ms.prod: outlook
ms.assetid: 430d453a-6727-4dc6-bc77-0ec9866b4a38
ms.date: 06/08/2017
---


# Page Members (Outlook Forms Script)

Represents one page of a  [MultiPage](Outlook.multipage.md) or a single member of a [Pages](pages-object-outlook-forms-script.md) collection.


## Methods



|**Name**|**Description**|
|:-----|:-----|
| [Copy](Outllok.Page.copy.md)|Copies the contents of an object to the Clipboard.|
| [Cut](Outllok.Page.cut.md)|Removes selected information from an object and transfers it to the Clipboard.|
| [Paste](Outllok.Page.paste.md)|Transfers the contents of the Clipboard to an object.|
| [RedoAction](Outllok.Page.redoaction.md)|Reverses the effect of the most recent  **Undo** action.|
| [Repaint](Outllok.Page.repaint.md)|Updates the display by redrawing the page.|
| [Scroll](Outllok.Page.scroll.md)|Moves the scroll bar on an object.|
| [SetDefaultTabOrder](Outllok.Page.setdefaulttaborder.md)|Sets the  **TabIndex** property of each control on a frame or page, using a default top-to-bottom, left-to-right tab order.|
| [UndoAction](Outllok.Page.undoaction.md)|Reverses the most recent action that supports the  **Undo** command.|



## Properties



|**Name**|**Description**|
|:-----|:-----|
| [Accelerator](Outllok.Page.accelerator.md)|Returns or sets the accelerator key for the page. Read/write.|
| [ActiveControl](Outllok.Page.activecontrol.md)|Returns an  **Object** that has the focus. Read-only.|
| [CanPaste](Outllok.Page.canpaste.md)|Returns a  **Boolean** that specifies whether the Clipboard contains data that the object supports. Read-only.|
| [CanRedo](Outllok.Page.canredo.md)|Returns a  **Boolean** that specifies if the most recent **Undo** can be reversed. Read-only.|
| [CanUndo](Outllok.Page.canundo.md)|Returns a  **Boolean** that specifies whether the last user action can be undone. Read-only.|
| [Caption](Outllok.Page.caption.md)|Returns or sets a  **String** that specifies the text that appears on the page. Read/write.|
| [ControlTipText](Outllok.Page.controltiptext.md)|Returns and sets a  **String** that specifies text that appears when the user briefly holds the mouse pointer over a control without clicking. Read/write.|
| [Cycle](Outllok.Page.cycle.md)|Returns or sets an  **Integer** that specifies whether cycling includes controls nested in a [MultiPage](Outlook.multipage.md). Read/write.|
| [Enabled](Outllok.Page.enabled.md)|Returns or sets a  **Boolean** that specifies whether a control can receive the focus and respond to user-generated events. Read/write.|
| [Index](Outllok.Page.index.md)|Returns or sets an  **Integer** that specifies the position of a [Page](Outllok.Page.object-outlook-forms-script.md) object in a [Pages](pages-object-outlook-forms-script.md) collection. Read/write.|
| [InsideHeight](Outllok.Page.insideheight.md)|Returns a  **Long** that specifies the height, in points, of the client region inside a **Page**. Read-only.|
| [InsideWidth](Outllok.Page.insidewidth.md)|Returns a  **Long** that specifies the width, in points, of the client region inside a **Page**. Read-only.|
| [KeepScrollBarsVisible](Outllok.Page.keepscrollbarsvisible.md)|Returns or sets an  **Integer** that specifies whether scroll bars remain visible when not required. Read/write.|
| [Name](Outllok.Page.name.md)|Returns or sets a  **String** that specifies the name of an object. Read/write.|
| [Parent](Outllok.Page.parent.md)|Returns an  **Object** that represents the parent object of the specified page. Read-only.|
| [Picture](Outllok.Page.picture.md)|Returns a  **String** that specifies the full path name of a bitmap to display on a control. Read-only.|
| [PictureAlignment](Outllok.Page.picturealignment.md)|Returns or sets an  **Integer** that specifies the location of a background picture. Read/write.|
| [PictureSizeMode](Outllok.Page.picturesizemode.md)|Returns or sets an  **Integer** that specifies how to display the background picture on a page. Read/write.|
| [PictureTiling](Outllok.Page.picturetiling.md)|Returns or sets a  **Boolean** that specifies whether a picture is repeated across the background of the object. Read/write.|
| [ScrollBars](Outllok.Page.scrollbars.md)|Returns or sets an  **Integer** that specifies whether a page has vertical scroll bars, horizontal scroll bars, or both. Read/write.|
| [ScrollHeight](Outllok.Page.scrollheight.md)|Returns or sets a  **Single** that specifies the height, in points, of the total area that can be viewed by moving the scroll bars on the **Page**. Read/write.|
| [ScrollLeft](Outllok.Page.scrollleft.md)|Returns or sets a  **Single** that specifies the distance, in points, of the left edge of the visible form from the left edge of the **Page**. Read/write.|
| [ScrollTop](Outllok.Page.scrolltop.md)|Returns or sets a  **Single** that specifies the distance, in points, of the top edge of the visible form from the top edge of the **Page**. Read/write.|
| [ScrollWidth](Outllok.Page.scrollwidth.md)|Returns or sets a  **Single** that specifies the width, in points, of the total area that can be viewed by moving the scroll bars on the **Page**. Read/write.|
| [Tag](Outllok.Page.tag.md)|Returns or sets a  **String** that specifies additional information about an object. Read/write.|
| [VerticalScrollBarSide](Outllok.Page.verticalscrollbarside.md)|Returns or sets an  **Integer** that specifies whether a vertical scroll bar appears on the right or left side of a page. Read/write.|
| [Visible](Outllok.Page.visible.md)|Returns or sets a  **Boolean** that specifies whether a **Page** is visible or hidden. Read/write.|
| [Zoom](Outllok.Page.zoom.md)|Returns or sets an  **Integer** that specifies the percentage to increase or decrease the displayed image. Read/write.|



