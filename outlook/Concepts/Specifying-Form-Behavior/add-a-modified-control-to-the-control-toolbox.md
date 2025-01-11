---
title: Add a Modified Control to the Control Toolbox
ms.assetid: 74d751d0-e93d-557e-e878-1fce71b7143a
ms.date: 06/08/2019
ms.localizationpriority: medium
---


# Add a Modified Control to the Control Toolbox

1. In an open form on the **Developer** tab, in the **Tools** group, select **Control Toolbox**.
![Control Toolbox icon](../../../images/0548_ZA06045100.gif)
**Note** If you don't see the **Developer** tab in the open form, see the [Run in Developer Mode in Outlook](../../How-to/Using-Visual-Basic-to-Customize-Outlook-Forms/run-in-developer-mode-in-outlook.md).
1. Drag a control from the **Control Toolbox** to your form and customize it. For example, to create an **OK** button, drag a **CommandButton** control onto the form, set its **Caption** property to **OK**, and set its **Default** property to **True**.
1. Select the customized control.
1. Drag the control to the **Control Toolbox**.

**Note** When you drag a control onto the **Control Toolbox**, you only transfer the advanced property values. Any lines of code or Outlook property values that you have written for that control don't transfer with the control. You must write new code or copy code from the control on the form to the control on the **Control Toolbox**.

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
