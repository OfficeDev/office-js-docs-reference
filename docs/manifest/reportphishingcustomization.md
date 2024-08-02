---
title: ReportPhishingCustomization element in the manifest file
description: The ReportPhishingCustomization element configures the ribbon button and preprocessing dialog of a spam-reporting add-in in Outlook.
ms.date: 05/20/2024
ms.localizationpriority: medium
---

# ReportPhishingCustomization element

Configures the ribbon button and preprocessing dialog of a spam-reporting add-in in Outlook.

To learn more about how to implement the spam reporting feature in your add-in, see [Implement an integrated spam-reporting add-in](/office/dev/add-ins/outlook/spam-reporting).

**Add-in type**: Mail

**Valid only in these VersionOverrides schemas**:

- Mail 1.1

For more information, see [Version overrides in the add-in only manifest](/office/dev/add-ins/develop/xml-manifest-overview#version-overrides-in-the-manifest).

**Associated with these requirement sets**:

- [Mailbox 1.14](../requirement-sets/outlook/requirement-set-1.14/outlook-requirement-set-1.14.md)

## Contained in

- **\<ExtensionPoint\>** element with the **xsi:type** attribute set to [ReportPhishingCommandSurface](extensionpoint.md#reportphishingcommandsurface).

## Attributes

None.

## Child elements

| Element | Required | Description |
| :------ | :------: | :------ |
| [Control](control.md) | Yes | Configures and adds the add-in button to the ribbon. The **xsi:type** attribute must be set to `Button` and the **xsi:type** attribute of its **\<Action\>** child element must be set to `ExecuteFunction`. |
| [PreProcessingDialog](preprocessingdialog.md) | Yes | Configures the preprocessing dialog shown after the add-in button is selected from the ribbon. This dialog allows users to provide additional information about a message they're reporting. |
| [SourceLocation element (version overrides)](customfunctionssourcelocation.md) | Yes | Specifies the location of the source JavaScript file. |

## Example

```xml
<ExtensionPoint xsi:type="ReportPhishingCommandSurface">
  <ReportPhishingCustomization>
    <!-- Configures the ribbon button. -->
    <Control xsi:type="Button" id="ReportingButton">
      <Label resid="ReportingButton.Label"/>
      <Supertip>
        <Title resid="ReportingButton.Label"/>
        <Description resid="ReportingButton.Description"/>
      </Supertip>
      <Icon>
        <bt:Image size="16" resid="Icon.16x16"/>
        <bt:Image size="32" resid="Icon.32x32"/>
        <bt:Image size="64" resid="Icon.64x64"/>
        <bt:Image size="80" resid="Icon.80x80"/>
      </Icon>
      <Action xsi:type="ExecuteFunction">
        <FunctionName>onMessageReport</FunctionName>
      </Action>
    </Control>
    <!-- Configures the preprocessing dialog. -->
    <PreProcessingDialog>
      <Title resid="PreProcessingDialog.Label"/>
      <Description resid="PreProcessingDialog.Description"/>
      <ReportingOptions>
        <Title resid="OptionsTitle.Label"/>
        <Option resid="Option1.Label"/>
        <Option resid="Option2.Label"/>
        <Option resid="Option3.Label"/>
        <Option resid="Option4.Label"/>
      </ReportingOptions>
      <FreeTextLabel resid="FreeText.Label"/>
      <MoreInfo>
        <MoreInfoText resid="MoreInfo.Label"/>
        <MoreInfoUrl resid="MoreInfo.Url"/>
      </MoreInfo>
    </PreProcessingDialog>
    <SourceLocation resid="Commands.Url"/>
  </ReportPhishingCustomization>
</ExtensionPoint>
```
