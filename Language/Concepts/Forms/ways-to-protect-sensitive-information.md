---
title: Ways to protect sensitive information
keywords: fm20.chm5225235
f1_keywords:
- fm20.chm5225235
ms.prod: office
ms.assetid: efd37fb2-7bec-d824-08cb-f8e50df40dd5
ms.date: 06/08/2017
---


# Ways to protect sensitive information

Many applications use data that should be available only to certain users. Here are some suggestions you can use to protect sensitive information in Microsoft Forms:



- Write code that makes a control (and its data) invisible to unauthorized users. The  **Visible** property makes a control visible or invisible. For more information about **Visible**, see[Visible Property](../../reference/User-Interface-Help/visible-property-microsoft-forms.md).
    
- Write code that sets the control's foreground and background to the same color when unauthorized users run the application. This hides the information from unauthorized users. The  **ForeColor** and **BackColor** properties determine the[foreground color](../../Glossary/glossary-vba.md) and the[background color](../../Glossary/glossary-vba.md). For information about  **ForeColor**, see[ForeColor Property](../../reference/User-Interface-Help/forecolor-property-microsoft-forms.md). For information about  **BackColor**, see[BackColor Property](../../reference/User-Interface-Help/backcolor-property-microsoft-forms.md).
    
- Disable the control when unauthorized users run the application. The  **Enabled** property determines when a control is disabled. For information about **Enabled**, see[Enabled Property](../../reference/User-Interface-Help/enabled-property-microsoft-forms.md).
    
- Require a password for access to the application or a specific control. You can use [placeholders](../../Glossary/glossary-vba.md) as the user types each character. The **PasswordChar** property defines placeholder characters. For information about **PasswordChar**, see[PasswordChar Property](passwordchar-property.md).
    


 **Note**  Using passwords or any other techniques listed can improve the security of your application, but does not guarantee the prevention of unauthorized access to your data.


