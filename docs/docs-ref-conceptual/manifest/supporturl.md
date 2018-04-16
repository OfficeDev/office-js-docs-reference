# SupportUrl element

Specifies the URL of a page that provides support information for your add-in.

## Attributes

|**Attribute**|**Type**|**Required**|**Description**|
|:-----|:-----|:-----|:-----|
|DefaultValue|URL|required|Specifies the default value for this setting, expressed for the locale specified in the [DefaultLocale](defaultlocale.md) element.|

## Child elements

|  Element | Required | Description  |
|:-----|:-----|:-----|
|  [Override](override.md)   | No | Specifies the setting for additional locale urls |

## Parent element

[OfficeApp](officeapp.md)

## Example

```XML
<OfficeApp>
...
  <IconUrl DefaultValue="https://contoso.com/assets/icon-32.png" />
  <HighResolutionIconUrl DefaultValue="https://contoso.com/assets/hi-res-icon.png"/>
  
  
  <SupportUrl DefaultValue="https://contoso.com/support " />
  
  
  <AppDomains>
  ...
  </AppDomains>
...
</OfficeApp>
```

