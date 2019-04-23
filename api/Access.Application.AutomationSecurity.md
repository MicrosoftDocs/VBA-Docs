---
title: Application.AutomationSecurity property (Access)
keywords: vbaac10.chm12611
f1_keywords:
- vbaac10.chm12611
ms.prod: access
api_name:
- Access.Application.AutomationSecurity
ms.assetid: 4589f050-4b0c-8dba-309a-98ad3921baa7
ms.date: 02/05/2019
localization_priority: Normal
---


# Application.AutomationSecurity property (Access)

Returns or sets an **[MsoAutomationSecurity](office.msoautomationsecurity.md)** constant that represents the security mode that Microsoft Access uses when it is programmatically opening files. Read/write.


## Syntax

_expression_.**AutomationSecurity**

_expression_ A variable that represents an **[Application](Access.Application.md)** object.


## Remarks

**MsoAutomationSecurity** can be one of these **MsoAutomationSecurity** constants:

- **msoAutomationSecurityByUI**. Uses the security setting specified in the **Security** dialog box (**Tools** menu > **Macro** submenu). **msoAutomationSecurityByUI** is the default value when the application is started.

- **msoAutomationSecurityForceDisable**. Access will not open any database if the macro security level is set to **High** or **Medium** in the **Security** dialog box (**Tools** menu > **Macro** submenu). No security messages are shown to the user.

  > [!NOTE] 
  > This setting has no effect if the macro security level is set to **Low**.

- **msoAutomationSecurityLow** Enables all macros.



[!include[Support and feedback](~/includes/feedback-boilerplate.md)]