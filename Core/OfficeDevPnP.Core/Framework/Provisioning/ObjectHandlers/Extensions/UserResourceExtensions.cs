using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Resources;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Extensions
{


    internal static class UserResourceExtensions
    {
        public static ProvisioningTemplate SaveResourceValues(ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            var tempFolder = System.IO.Path.GetTempPath();

            var languages = new List<int>(creationInfo.ResourceTokens.Keys.Select(t => t.Item2).Distinct());
            foreach (int language in languages)
            {
                var culture = new CultureInfo(language);

                var resourceFileName = System.IO.Path.Combine(tempFolder, $"{creationInfo.ResourceFilePrefix}.{culture.Name}.resx");
                if (System.IO.File.Exists(resourceFileName))
                {
                    // Read existing entries, if any
#if !NETSTANDARD2_0
                    using (ResXResourceReader resxReader = new ResXResourceReader(resourceFileName))
#else
                    using (ResourceReader resxReader = new ResourceReader(resourceFileName))
#endif
                    {
                        foreach (DictionaryEntry entry in resxReader)
                        {
                            // find if token is already there
                            if (!creationInfo.ResourceTokens.ContainsKey(new Tuple<string, int>(entry.Key.ToString(), language)))
                            {
                                creationInfo.ResourceTokens.Add(new Tuple<string, int>(entry.Key.ToString(), language), entry.Value as string);
                            }
                        }
                    }
                }

                // Create new resource file
#if !NETSTANDARD2_0
                using (ResXResourceWriter resx = new ResXResourceWriter(resourceFileName))
#else
                using (ResourceWriter resx = new ResourceWriter(resourceFileName))
#endif
                {
                    foreach (var token in creationInfo.ResourceTokens.Where(t => t.Key.Item2 == language))
                    {
                        resx.AddResource(token.Key.Item1, token.Value);
                    }
                }

                template.Localizations.Add(new Localization() { LCID = language, Name = culture.NativeName, ResourceFile = $"{creationInfo.ResourceFilePrefix}.{culture.Name}.resx" });

                // Persist the file using the connector
                using (FileStream stream = System.IO.File.Open(resourceFileName, FileMode.Open))
                {
                    creationInfo.FileConnector.SaveFileStream($"{creationInfo.ResourceFilePrefix}.{culture.Name}.resx", stream);
                }
                // remove the temp resx file
                System.IO.File.Delete(resourceFileName);
            }
            return template;
        }

        public static bool SetUserResourceValue(this UserResource userResource, string tokenValue, TokenParser parser)
        {
            bool isDirty = false;

            if (userResource != null && !String.IsNullOrEmpty(tokenValue))
            {
                var resourceValues = parser.GetResourceTokenResourceValues(tokenValue);
                foreach (var resourceValue in resourceValues)
                {
                    userResource.SetValueForUICulture(resourceValue.Item1, resourceValue.Item2);
                    isDirty = true;
                }
            }

            return isDirty;
        }

        public static bool PersistResourceValue(UserResource userResource, string token, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
			(userResource.Context as ClientContext).Web.EnsureProperty(w => w.Language);

			bool returnValue = false;
            foreach (var language in template.SupportedUILanguages)
            {
				if (language.LCID == (userResource.Context as ClientContext).Web.Language) //Ignore default language
					continue;

                var culture = new CultureInfo(language.LCID);

                var value = userResource.GetValueForUICulture(culture.Name);
                userResource.Context.ExecuteQueryRetry();
                if (!string.IsNullOrEmpty(value.Value))
                {
                    returnValue = true;

                    if (!creationInfo.ResourceTokens.ContainsKey(new Tuple<string, int>(token, language.LCID)))
                        creationInfo.ResourceTokens.Add(new Tuple<string, int>(token, language.LCID), value.Value);
                }
            }

            return returnValue;
        }

        public static bool PersistResourceValue(string token, int LCID, string Title, ProvisioningTemplateCreationInformation creationInfo)
        {
            bool returnValue = false;

            if (!string.IsNullOrWhiteSpace(Title))
            {
                returnValue = true;

                if (!creationInfo.ResourceTokens.ContainsKey(new Tuple<string, int>(token, LCID)))
                    creationInfo.ResourceTokens.Add(new Tuple<string, int>(token, LCID), Title);
            }

            return returnValue;
        }

        public static bool PersistResourceValue(List siteList, Guid viewId, string token, ProvisioningTemplate template, ProvisioningTemplateCreationInformation creationInfo)
        {
            bool returnValue = false;
            var clientContext = siteList.Context;

            foreach (var language in template.SupportedUILanguages)
            {
                var culture = new CultureInfo(language.LCID);
                var currentView = siteList.GetViewById(viewId);
                clientContext.Load(currentView, cc => cc.Title);
                var acceptLanguage = clientContext.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"];
                clientContext.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = new CultureInfo(language.LCID).Name;
                clientContext.ExecuteQueryRetry();

                if (!string.IsNullOrWhiteSpace(currentView.Title))
                {
                    returnValue = true;

                    if (!creationInfo.ResourceTokens.ContainsKey(new Tuple<string, int>(token, language.LCID)))
                        creationInfo.ResourceTokens.Add(new Tuple<string, int>(token, language.LCID), currentView.Title);
                }

                clientContext.PendingRequest.RequestExecutor.WebRequest.Headers["Accept-Language"] = acceptLanguage;

            }
            return returnValue;
        }

        public static bool ContainsResourceToken(this string value)
        {
            if (value != null)
            {
                return Regex.IsMatch(value, "\\{(res|loc|resource|localize|localization):(.*?)(\\})", RegexOptions.IgnoreCase);
            }
            else
            {
                return false;
            }

        }
    }

}
