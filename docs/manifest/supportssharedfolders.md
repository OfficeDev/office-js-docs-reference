---
title: SupportsSharedFolders element in the manifest file
description: The SupportsSharedFolders element defines whether the Outlook add-in is available in shared folders and shared mailbox scenarios.
ms.date: 05/19/2023
ms.localizationpriority: medium
---

# SupportsSharedFolders element

Defines whether the Outlook add-in is available in shared folders (that is, delegate access) and shared mailboxes scenarios. The **\<SupportsSharedFolders\>** element is a child element of [DesktopFormFactor](desktopformfactor.md). It is set to *false* by default.

To learn more about shared folder and shared mailbox scenarios, see [Enable shared folders and shared mailbox scenarios in an Outlook add-in](/office/dev/add-ins/outlook/delegate-access).

> [!IMPORTANT]
> Support for this element in shared folder scenarios was introduced in requirement set 1.8, while support in shared mailbox scenarios was introduced in requirement set 1.13. See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support these requirement sets.

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.8 for shared folder scenarios](../requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8.md)
- [Mailbox 1.13 for shared mailbox scenarios](../requirement-sets/outlook/requirement-set-1.13/outlook-requirement-set-1.13.md)

The following is an example of the **\<SupportsSharedFolders\>** element.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- Configure selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```
