---
title: Use z-order to layer controls
keywords: fm20.chm5225259
f1_keywords:
- fm20.chm5225259
ms.prod: office
ms.assetid: 07357aa8-bcd0-3ad0-a4e3-c059b5f17b7d
ms.date: 12/29/2018
localization_priority: Normal
---


# Use z-order to layer controls

To place a control at the front or back of the [z-order](../../Glossary/vbe-glossary.md#z-order):

1. Select the controls that you want to reposition.
    
2. From the **[Format](../../reference/user-interface-help/format-menu.md)** menu, choose **Order**.
    
3. From the cascading menu, select **Bring to Front** or **Send to Back**.
    

To adjust a control one position in the z-order:

1. Select the controls that you want to reposition.
    
2. From the **Format** menu, choose **Order**.
    
3. From the cascading menu, select **Bring Forward** or **Send Backward**.
    
> [!NOTE] 
> You can't undo or redo layering commands, such as **Send to Back** or **Bring to Front**. For example, if you select an object and choose **Send Backward** on the shortcut menu, you won't be able to undo or redo that action.

The **Bring to Front**, **Bring Forward**, **Send to Back**, and **Send Backward** menu choices let you change the z-order of a control relative to other controls. If the form includes any **[ListBox](../../reference/user-interface-help/listbox-control.md)**, **[Frame](../../reference/user-interface-help/frame-control.md)**, or **[MultiPage](../../reference/user-interface-help/multipage-control.md)** controls, those controls automatically move as close as possible to the top of the stack. For example, applying **Send Backward** to a **ListBox**, **Frame**, or **MultiPage** moves the control below other **ListBox**, **Frame**, or **MultiPage** controls, but will not move it below any other type of control in the stack. 

Similarly, applying **Bring Forward** to a control other than a **ListBox**, **Frame**, or **MultiPage** will move the control closer to the top of the stack, but will not move it above any **ListBox**, **Frame**, or **MultiPage** in the stack.

Visually, this means that if a **ListBox**, **Frame**, or **MultiPage** and any other Microsoft Forms control are in the same location on a form, the **ListBox**, **Frame**, or **MultiPage** will always appear on top of the other control. If a **ListBox**, **Frame**, or **MultiPage** is in the same place as another **ListBox**, **Frame**, or **MultiPage**, the z-order of the controls determines which control appears on top of the other.

## See also

- [Microsoft Forms collections, controls, and objects](../../reference/user-interface-help/objects-microsoft-forms.md)
- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]