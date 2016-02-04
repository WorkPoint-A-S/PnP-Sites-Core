﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using OfficeDevPnP.Core.Enums;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Field = OfficeDevPnP.Core.Framework.Provisioning.Model.Field;
using SPField = Microsoft.SharePoint.Client.Field;
using OfficeDevPnP.Core.Diagnostics;
using System.Text.RegularExpressions;
using System.Globalization;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    internal class ObjectField : ObjectHandlerBase
    {
        public override string Name
        {
            get { return "Fields"; }
        }
        public override TokenParser ProvisionObjects(Web web, ProvisioningTemplate template, TokenParser parser, ProvisioningTemplateApplyingInformation applyingInformation)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not provisioning fields. Technically this can be done but it's not a recommended practice
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Context_web_is_subweb__skipping_site_columns);
                    return parser;
                }

                var existingFields = web.Fields;

                web.Context.Load(web, w => w.RegionalSettings.LocaleId);
                web.Context.Load(existingFields, fs => fs.Include(f => f.Id, f => f.TitleResource, f => f.DescriptionResource));
                web.Context.ExecuteQueryRetry();
                var existingFieldIds = existingFields.AsEnumerable<SPField>().Select(l => l.Id).ToList();
                var fields = template.SiteFields;

                var cultureNames = new List<string>();
                var webCultureInfo = new CultureInfo((int)web.RegionalSettings.LocaleId);

                foreach (var supportedUILanguage in template.SupportedUILanguages)
                {
                    var ci = new CultureInfo(supportedUILanguage.LCID);
                    cultureNames.Add(ci.Name);
                }

                foreach (var field in fields)
                {
                    XElement templateFieldElement = XElement.Parse(parser.ParseString(field.SchemaXml, "~sitecollection", "~site"));
                    var fieldId = templateFieldElement.Attribute("ID").Value;
                    SPField provisionedField = null;

                    if (!existingFieldIds.Contains(Guid.Parse(fieldId)))
                    {
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Adding_field__0__to_site, fieldId);
                            provisionedField = CreateField(web, templateFieldElement, scope, parser);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Fields_Adding_field__0__failed___1_____2_, fieldId, ex.Message, ex.StackTrace);
                            throw;
                        }
                    }
                    else
                        try
                        {
                            scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__in_site, fieldId);
                            provisionedField = UpdateField(web, fieldId, templateFieldElement, scope, parser);
                        }
                        catch (Exception ex)
                        {
                            scope.LogError(CoreResources.Provisioning_ObjectHandlers_Fields_Updating_field__0__failed___1_____2_, fieldId, ex.Message, ex.StackTrace);
                            throw;
                        }

                    // Localizations
                    if (provisionedField != null)
                    {
                        provisionedField.EnsureProperties(f => f.Title, f => f.Description);

                        // IMPORTANT! Title or Description corresponding to the culture of the web, must be set first, otherwise localizations doesn't work.
                        var primaryLocalization = template.SiteFieldsLocalizations.FirstOrDefault(l => l.Id.Equals(Guid.Parse(fieldId)) && webCultureInfo.Name.Equals(l.CultureName, StringComparison.InvariantCultureIgnoreCase));
                        if (primaryLocalization != null && (provisionedField.Title != primaryLocalization.TitleResource || provisionedField.Description != primaryLocalization.DescriptionResource))
                        {
                            provisionedField.Title = primaryLocalization.TitleResource;
                            provisionedField.Description = primaryLocalization.DescriptionResource;
                }

                        foreach (var localization in template.SiteFieldsLocalizations.Where(l => l.Id.Equals(Guid.Parse(fieldId)) && cultureNames.Contains(l.CultureName, StringComparer.InvariantCultureIgnoreCase)))
                        {
                            provisionedField.SetLocalizationForField(localization.CultureName, localization.TitleResource, localization.DescriptionResource);
            }
                    }
                }
            }
            return parser;
        }

        private SPField UpdateField(Web web, string fieldId, XElement templateFieldElement, PnPMonitoredScope scope, TokenParser parser)
        {
            var existingField = web.Fields.GetById(Guid.Parse(fieldId));
            web.Context.Load(existingField, f => f.SchemaXml);
            web.Context.ExecuteQueryRetry();

            XElement existingFieldElement = XElement.Parse(existingField.SchemaXml);

            XNodeEqualityComparer equalityComparer = new XNodeEqualityComparer();

            if (equalityComparer.GetHashCode(existingFieldElement) != equalityComparer.GetHashCode(templateFieldElement)) // Is field different in template?
            {
                if (existingFieldElement.Attribute("Type").Value == templateFieldElement.Attribute("Type").Value) // Is existing field of the same type?
                {
                    var listIdentifier = templateFieldElement.Attribute("List") != null ? templateFieldElement.Attribute("List").Value : null;

                    if (listIdentifier != null)
                    {
                        // Temporary remove list attribute from list
                        templateFieldElement.Attribute("List").Remove();
                    }

                    foreach (var attribute in templateFieldElement.Attributes())
                    {
                        if (existingFieldElement.Attribute(attribute.Name) != null)
                        {
                            existingFieldElement.Attribute(attribute.Name).Value = attribute.Value;
                        }
                        else
                        {
                            existingFieldElement.Add(attribute);
                        }
                    }
                    foreach (var element in templateFieldElement.Elements())
                    {
                        if (existingFieldElement.Element(element.Name) != null)
                        {
                            existingFieldElement.Element(element.Name).Remove();
                        }
                        existingFieldElement.Add(element);
                    }

                    if (existingFieldElement.Attribute("Version") != null)
                    {
                        existingFieldElement.Attributes("Version").Remove();
                    }
                    existingField.SchemaXml = parser.ParseString(existingFieldElement.ToString(), "~sitecollection", "~site");
                    existingField.UpdateAndPushChanges(true);
                    web.Context.ExecuteQueryRetry();
                }
                else
                {
                    var fieldName = existingFieldElement.Attribute("Name") != null ? existingFieldElement.Attribute("Name").Value : existingFieldElement.Attribute("StaticName").Value;
                    WriteWarning(string.Format(CoreResources.Provisioning_ObjectHandlers_Fields_Field__0____1___exists_but_is_of_different_type__Skipping_field_, fieldName, fieldId), ProvisioningMessageType.Warning);
                    scope.LogWarning(CoreResources.Provisioning_ObjectHandlers_Fields_Field__0____1___exists_but_is_of_different_type__Skipping_field_, fieldName, fieldId);
                }
            }

            return existingField;
        }

        private string ParseFieldSchema(string schemaXml, ListCollection lists)
        {
            foreach (var list in lists)
            {
                schemaXml = Regex.Replace(schemaXml, list.Id.ToString(), string.Format("{{listid:{0}}}", list.Title), RegexOptions.IgnoreCase);
            }

            return schemaXml;
        }

        private static SPField CreateField(Web web, XElement templateFieldElement, PnPMonitoredScope scope, TokenParser parser)
        {
            var listIdentifier = templateFieldElement.Attribute("List") != null ? templateFieldElement.Attribute("List").Value : null;

            if (listIdentifier != null)
            {
                // Temporary remove list attribute from list
                templateFieldElement.Attribute("List").Remove();
            }

            var fieldXml = parser.ParseString(templateFieldElement.ToString(), "~sitecollection", "~site");

            var field = web.Fields.AddFieldAsXml(fieldXml, false, AddFieldOptions.AddFieldInternalNameHint);
            web.Context.ExecuteQueryRetry();

            return field;
        }

        public override ProvisioningTemplate ExtractObjects(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            using (var scope = new PnPMonitoredScope(this.Name))
            {
                // if this is a sub site then we're not creating field entities.
                if (web.IsSubSite())
                {
                    scope.LogDebug(CoreResources.Provisioning_ObjectHandlers_Fields_Context_web_is_subweb__skipping_site_columns);
                    return template;
                }

                var existingFields = web.Fields;
                web.Context.Load(web, w => w.ServerRelativeUrl, w => w.IsMultilingual, w => w.SupportedUILanguageIds);
                web.Context.Load(existingFields, fs => fs.Include(f => f.Id, f => f.SchemaXml, f => f.TypeAsString, f => f.TitleResource, f => f.DescriptionResource));
                web.Context.Load(web.Lists, ls => ls.Include(l => l.Id, l => l.Title));
                web.Context.ExecuteQueryRetry();

                var taxTextFieldsToMoveUp = new List<Guid>();
                var cultureNames = new List<string>();

                if (web.IsMultilingual)
                {
                    foreach (var supportedlanguageId in web.SupportedUILanguageIds)
                    {
                        var ci = new CultureInfo(supportedlanguageId);
                        cultureNames.Add(ci.Name);
                    }
                }

                foreach (var field in existingFields)
                {
                    if (!BuiltInFieldId.Contains(field.Id))
                    {
                        var fieldXml = field.SchemaXml;
                        XElement element = XElement.Parse(fieldXml);

                        // Check if the field contains a reference to a list. If by Guid, rewrite the value of the attribute to use web relative paths
                        var listIdentifier = element.Attribute("List") != null ? element.Attribute("List").Value : null;
                        if (!string.IsNullOrEmpty(listIdentifier))
                        {
                            var listGuid = Guid.Empty;
                            fieldXml = ParseFieldSchema(fieldXml, web.Lists);
                            element = XElement.Parse(fieldXml);
                            //if (Guid.TryParse(listIdentifier, out listGuid))
                            //{
                            //    fieldXml = ParseListSchema(fieldXml, web.Lists);
                                //if (newfieldXml == fieldXml)
                                //{
                                //    var list = web.Lists.GetById(listGuid);
                                //    web.Context.Load(list, l => l.RootFolder.ServerRelativeUrl);
                                //    web.Context.ExecuteQueryRetry();

                                //    var listUrl = list.RootFolder.ServerRelativeUrl.Substring(web.ServerRelativeUrl.Length).TrimStart('/');
                                //    element.Attribute("List").SetValue(listUrl);
                                //    fieldXml = element.ToString();
                                //}
                            //}
                        }
                        // Check if the field is of type TaxonomyField
                        if (field.TypeAsString.StartsWith("TaxonomyField"))
                        {
                            var taxField = (TaxonomyField)field;
                            web.Context.Load(taxField, tf => tf.TextField, tf => tf.Id);
                            web.Context.ExecuteQueryRetry();
                            taxTextFieldsToMoveUp.Add(taxField.TextField);
                        }
                        // Check if we have version attribute. Remove if exists 
                        if (element.Attribute("Version") != null)
                        {
                            element.Attributes("Version").Remove();
                            fieldXml = element.ToString();
                        }
                        template.SiteFields.Add(new Field() { SchemaXml = fieldXml });

                        // Localizations
                        // Don't extract localization for Site Columns in the base template

                        if (creationInfo.BaseTemplate != null && !creationInfo.BaseTemplate.SiteFields.Any(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(field.Id)))
                        {
                            foreach (var cultureName in cultureNames)
                            {
                                var titleResource = field.TitleResource.GetValueForUICulture(cultureName);
                                var descriptionResource = field.DescriptionResource.GetValueForUICulture(cultureName);
                                field.Context.ExecuteQueryRetry();

                                if (!string.IsNullOrEmpty(titleResource.Value) || !string.IsNullOrEmpty(descriptionResource.Value))
                                {
                                    template.SiteFieldsLocalizations.Add(new Localization(cultureName)
                                    {
                                        Id = field.Id,
                                        TitleResource = titleResource.Value,
                                        DescriptionResource = descriptionResource.Value
                                    });
                                }
                            }
                        }
                    }
                }
                // move hidden taxonomy text fields to the top of the list
                foreach (var textFieldId in taxTextFieldsToMoveUp)
                {
                    var field = template.SiteFields.First(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(textFieldId));
                    template.SiteFields.RemoveAll(f => Guid.Parse(f.SchemaXml.ElementAttributeValue("ID")).Equals(textFieldId));
                    template.SiteFields.Insert(0, field);
                }
                // If a base template is specified then use that one to "cleanup" the generated template model
                if (creationInfo.BaseTemplate != null)
                {
                    template = CleanupEntities(template, creationInfo.BaseTemplate);
                }
            }
            return template;
        }

        private ProvisioningTemplate CleanupEntities(ProvisioningTemplate template, ProvisioningTemplate baseTemplate)
        {
            foreach (var field in baseTemplate.SiteFields)
            {

                XDocument xDoc = XDocument.Parse(field.SchemaXml);
                var id = xDoc.Root.Attribute("ID") != null ? xDoc.Root.Attribute("ID").Value : null;
                if (id != null)
                {
                    int index = template.SiteFields.FindIndex(f => f.SchemaXml.IndexOf(id, StringComparison.InvariantCultureIgnoreCase) > -1);

                    if (index > -1)
                    {
                        template.SiteFields.RemoveAt(index);
                    }
                }
            }

            return template;
        }

        public override bool WillProvision(Web web, ProvisioningTemplate template)
        {
            if (!_willProvision.HasValue)
            {
                _willProvision = template.SiteFields.Any();
            }
            return _willProvision.Value;
        }

        public override bool WillExtract(Web web, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            if (!_willExtract.HasValue)
            {
                _willExtract = true;
            }
            return _willExtract.Value;
        }
    }

    internal static class XElementStringExtensions
    {
        public static string ElementAttributeValue(this string input, string attribute)
        {
            var element = XElement.Parse(input);
            return element.Attribute(attribute) != null ? element.Attribute(attribute).Value : null;
        }
    }
}
