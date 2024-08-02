---
title: ExtendedPermission element in the manifest file
description: Defines an extended permission the add-in needs to access the associated API or feature.
ms.date: 05/19/2023
ms.localizationpriority: medium
---

# `ExtendedPermission` element

Defines an extended permission the add-in needs to access the associated API or feature. The `ExtendedPermission` element is a child element of [ExtendedPermissions](extendedpermissions.md).

> [!IMPORTANT]
> Support for this element was introduced in requirement set 1.9. See [clients and platforms](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients) that support this requirement set.

**Add-in type:** Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.9](../requirement-sets/outlook/requirement-set-1.9/outlook-requirement-set-1.9.md)

## Available extended permissions

The following are the available values.

|Available value|Description|Hosts|
|:---|:---|:---|
|`AppendOnSend`|Declares that the add-in uses the [Office.Body.prependOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-prependonsendasync-member(1)) or [Office.Body.appendOnSendAsync](/javascript/api/outlook/office.body#outlook-office-body-appendonsendasync-member(1)) API.<br><br>To learn more, see [Prepend or append content to a message or appointment body on send](/office/dev/add-ins/outlook/append-on-send).|Outlook|

## `ExtendedPermission` example

The following is an example of the `ExtendedPermission` element.

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
            <!-- Configure your selected extension point. -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed. -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
    <!-- Configure the prepend-on-send or append-on-send feature using the AppendOnSend value. -->
    <ExtendedPermissions>
      <ExtendedPermission>AppendOnSend</ExtendedPermission>
    </ExtendedPermissions>
  </VersionOverrides>
</VersionOverrides>
...
```

## Contained in

- [ExtendedPermissions](extendedpermissions.md)
