---
title: Stop code execution (VBA)
keywords: vbhw6.chm1008937
f1_keywords:
- vbhw6.chm1008937
ms.prod: office
ms.assetid: f0608ca1-d6d8-d722-cbd7-8a31634264ed
ms.date: 12/27/2018
localization_priority: Priority
---


# Stop code execution

As you run your code, it may stop executing for one of the following reasons:

- An untrapped [run-time error](../Glossary/vbe-glossary.md#run-time-error) occurs.
    
- A trapped run-time error occurs, and **Break on All Errors** is selected on the **[General](../reference/user-interface-help/options-dialog-box.md#general-tab)** tab of the **Options** dialog box.
  
- A [breakpoint](../Glossary/vbe-glossary.md#breakpoint) is encountered.
    
- A **[Stop](../reference/user-interface-help/stop-statement.md)** statement in your code is encountered, switching the mode to [break mode](../Glossary/vbe-glossary.md#break-mode).
    
- An **[End](../reference/user-interface-help/end-statement.md)** statement in your code is encountered, switching the mode to [design time](../Glossary/vbe-glossary.md#design-time).
    
- You halt execution manually at a given point.
    
- A [watch expression](../Glossary/vbe-glossary.md#watch-expression) that you set to break if its value changes or becomes true is encountered.
    
**To halt execution manually**

- To switch to break mode, from the **[Run](../reference/user-interface-help/run-menu.md)** menu, choose **Break** (CTRL+BREAK), or use the toolbar shortcut: ![Toolbar button](../../images/tbr_brk_ZA01201682.gif).
    
- To switch to design time, from the **Run** menu, choose **Reset <projectname&gt;**, or use the toolbar shortcut: ![Toolbar button](../../images/tbr_end_ZA01201701.gif).
    
**To continue execution when your application has halted**

- From the **[Debug](../reference/user-interface-help/debug-menu.md)** menu, choose **Step Into** (F8), **Step Over** (SHIFT+F8), **Step Out** (CTRL+SHIFT+F8), or **Run To Cursor** (CTRL+F8).
    
## See also

- [Visual Basic how-to topics](../reference/user-interface-help/visual-basic-how-to-topics.md)

[!include[Support and feedback](~/includes/feedback-boilerplate.md)]
