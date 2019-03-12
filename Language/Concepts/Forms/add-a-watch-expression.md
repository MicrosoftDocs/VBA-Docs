---
title: Add a watch expression (VBA)
keywords: vbhw6.chm1008851
f1_keywords:
- vbhw6.chm1008851
ms.prod: office
ms.assetid: 0271930b-3238-ad36-f18f-1fdbc96ca766
ms.date: 12/27/2018
localization_priority: Normal
---


# Add a watch expression

A watch expression is an expression that you define to be monitored in the [Watch window](../../reference/user-interface-help/watch-window.md). When your application enters [break mode](../../Glossary/vbe-glossary.md#break-mode), the watch expressions you selected appear in the Watch window where you can observe their values.

**To add a watch expression**

1. On the **[Debug](../../reference/user-interface-help/debug-menu.md)** menu, choose **Add Watch**. The **[Add Watch](../../reference/user-interface-help/add-watch-dialog-box.md)** dialog box is displayed.
    
2. If an [expression](../../Glossary/vbe-glossary.md#expression) is already selected in the [Code window](../../reference/user-interface-help/code-window.md), it is automatically displayed in the **Expression** box. If no expression is displayed, enter the expression you want to evaluate. The expression can be a [variable](../../Glossary/vbe-glossary.md#variable), a [property](../../Glossary/vbe-glossary.md#property), a function call, or any other valid expression.
    
3. Select a [module](../../Glossary/vbe-glossary.md#module) or [procedure](../../Glossary/vbe-glossary.md#procedure) context in the **Context** group to select the range for which the expression will be evaluated.
    
   > [!NOTE] 
   > Select the narrowest [scope](../../Glossary/vbe-glossary.md#scope) that fits your needs. Selecting all procedures or all modules can slow down module execution considerably, because the expression is evaluated after execution of each statement. If you select a specific procedure for a context, execution is affected only while the procedure is in the list of active procedure calls. Choose **Call Stack** from the **[View](../../reference/user-interface-help/view-menu.md)** menu to display the list of active procedures.

4. Select an option in the **Watch Type** group to define how the system responds to the watch expression.
    
   - To display the value of the watch expression, choose **Watch Expression**.
    
   - To stop execution if the expression evaluates to **True**, choose **Break When Value is True**.
    
   - To stop execution when the value of the expression changes, choose **Break When Value Changes**.
     
5. Choose **OK**.
    
## See also

- [Visual Basic how-to topics](../../reference/user-interface-help/visual-basic-how-to-topics.md)
- [Microsoft Forms reference](../../reference/user-interface-help/reference-microsoft-forms.md)
- [Microsoft Forms conceptual topics](../../reference/user-interface-help/concepts-microsoft-forms.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
