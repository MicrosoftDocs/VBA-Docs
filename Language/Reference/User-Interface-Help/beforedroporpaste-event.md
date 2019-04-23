---
title: BeforeDropOrPaste event
keywords: fm20.chm5224936
f1_keywords:
- fm20.chm5224936
ms.prod: office
api_name:
- Office.BeforeDropOrPaste
ms.assetid: ba572265-1a9d-2d02-6346-82f88c1f249a
ms.date: 11/15/2018
localization_priority: Normal
---


# BeforeDropOrPaste event

Occurs when the user is about to drop or paste data onto an object.

## Syntax

For Frame  <br/>
**Private Sub**_object_ _**BeforeDropOrPaste( ByVal**_Cancel_**As MSForms.ReturnBoolean**, <br/>
_ctrl_**As Control**, <br/>
**ByVal**_Action_**As fmAction**, <br/>
**ByVal**_Data_**As DataObject**, <br/>
**ByVal**_X_**As Single**, <br/>
**ByVal**_Y_**As Single**, <br/>
**ByVal**_Effect_**As MSForms.ReturnEffect**, <br/>
**ByVal**_Shift_**As fmShiftState)**

For MultiPage  <br/>
**Private Sub**_object_ _**BeforeDropOrPaste(**_index_**As Long**, <br/>
**ByVal**_Cancel_**As MSForms.ReturnBoolean**, <br/>
_ctrl_**As Control**, <br/>
**ByVal**_Action_**As fmAction**, <br/>
**ByVal**_Data_**As DataObject**, <br/>
**ByVal**_X_**As Single**, <br/>
**ByVal**_Y_**As Single**, <br/>
**ByVal**_Effect_**As MSForms.ReturnEffect**, <br/>
**ByVal**_Shift_**As fmShiftState)**

For TabStrip  <br/>
**Private Sub**_object_ _**BeforeDropOrPaste(**_index_**As Long**, <br/>
**ByVal**_Cancel_**As MSForms.ReturnBoolean**, <br/>
**ByVal**_Action_**As fmAction**, <br/>
**ByVal**_Data_**As DataObject**, <br/>
**ByVal**_X_**As Single**, <br/>
**ByVal**_Y_**As Single**, <br/>
**ByVal**_Effect_**As MSForms.ReturnEffect**, <br/>
**ByVal**_Shift_**As fmShiftState)**

For other controls  <br/>
**Private Sub**_object_ _**BeforeDropOrPaste( ByVal**_Cancel_**As MSForms.ReturnBoolean**, <br/>
**ByVal**_Action_**As fmAction**, <br/>
**ByVal**_Data_**As DataObject**, <br/>
**ByVal**_X_**As Single**, <br/>
**ByVal**_Y_**As Single**, <br/>
**ByVal**_Effect_**As MSForms.ReturnEffect**, <br/>
**ByVal**_Shift_**As fmShiftState)**

The **BeforeDropOrPaste** event syntax has these parts:

|Part|Description|
|:-----|:-----|
| _object_|Required. A valid object name.|
| _index_|Required. The index of the **Page** in a **[MultiPage](multipage-control.md)** that the drop or paste operation will affect.|
| _Cancel_|Required. Event status. **False** indicates that the control should handle the event (default). **True** indicates the application handles the event.|
| _ctrl_|Required. The target control.|
| _Action_|Required. Indicates the result, based on the current keyboard settings, of the pending drag-and-drop operation.|
| _Data_|Required. Data that is dragged in a drag-and-drop operation. The data is packaged in a **[DataObject](dataobject-object.md)**.|
| _X, Y_|Required. The horizontal and vertical position of the mouse pointer when the drop occurs. Both coordinates are measured in points.  _X_ is measured from the left edge of the control; _Y_ is measured from the top of the control..|
| _Effect_|Required. Effect of the drag-and-drop operation on the target control.|
| _Shift_|Required. Specifies the state of SHIFT, CTRL, and ALT.|

## Settings

The settings for _Action_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmActionPaste_|2|Pastes the selected object into the drop target.|
| _fmActionDragDrop_|3|Indicates the user has dragged the object from its source to the drop target and dropped it on the drop target.|

<br/>

The settings for _Effect_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmDropEffectNone_|0|Does not copy or move the [drop source](../../Glossary/glossary-vba.md#drop-source) to the drop target.|
| _fmDropEffectCopy_|1|Copies the drop source to the drop target.|
| _fmDropEffectMove_|2|Moves the drop source to the drop target.|
| _fmDropEffectCopyOrMove_|3|Copies or moves the drop source to the drop target.|

<br/>

The settings for _Shift_ are:

|Constant|Value|Description|
|:-----|:-----|:-----|
| _fmShiftMask_|1|SHIFT was pressed.|
| _fmCtrlMask_|2|CTRL was pressed.|
| _fmAltMask_|4|ALT was pressed.|

## Remarks

For a **[MultiPage](multipage-control.md)** or **[TabStrip](tabstrip-control.md)**, Visual Basic for Applications initiates this event when it transfers a data object to the control.

For other controls, the system initiates this event immediately prior to the drop or paste operation.

When a control handles this event, you can update the _Action_ argument to identify the drag-and-drop action to perform. 

When _Effect_ is set to **fmDropEffectCopyOrMove**, you can assign _Action_ to **fmDropEffectNone**, **fmDropEffectCopy**, or **fmDropEffectMove**. 

When _Effect_ is set to **fmDropEffectCopy** or **fmDropEffectMove**, you can reassign _Action_ to **fmDropEffectNone**. You cannot reassign _Action_ when _Effect_ is set to **fmDropEffectNone**.

## See also

- [Microsoft Forms examples](examples-microsoft-forms.md)
- [Microsoft Forms reference](reference-microsoft-forms.md)
- [Microsoft Forms concepts](concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]