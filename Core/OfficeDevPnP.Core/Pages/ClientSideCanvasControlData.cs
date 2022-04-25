using Newtonsoft.Json;

namespace OfficeDevPnP.Core.Pages
{
#if !SP2013 && !SP2016
    /// <summary>
    /// Base class representing the json control data that will be included in each client side control (de-)serialization (data-sp-controldata attribute)
    /// </summary>
    public class ClientSideCanvasControlData
    {
        /// <summary>
        /// Gets or sets JsonProperty "controlType"
        /// </summary>
        [JsonProperty(PropertyName = "controlType")]
        public int ControlType { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "id"
        /// </summary>
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        /// <summary>
        /// Gets or sets JsonProperty "position"
        /// </summary>
        [JsonProperty(PropertyName = "position", NullValueHandling = NullValueHandling.Ignore)]
        public ClientSideCanvasControlPosition Position { get; set; }

        [JsonProperty(PropertyName = "emphasis", NullValueHandling = NullValueHandling.Ignore)]
        public ClientSideSectionEmphasis Emphasis { get; set; }
        [JsonProperty(PropertyName = "zoneGroupMetadata", NullValueHandling = NullValueHandling.Ignore)]
        public ZoneGroupMetadata ZoneGroupMetadata {  get; set; }
    }

    public class ZoneGroupMetadata
    {
        [JsonProperty(PropertyName = "type", NullValueHandling = NullValueHandling.Ignore)]
        public int Type { get; set; }
        [JsonProperty(PropertyName = "displayName", NullValueHandling = NullValueHandling.Ignore)]
        public string DisplayName { get; set; }
        [JsonProperty(PropertyName = "isExpanded", NullValueHandling = NullValueHandling.Ignore)]
        public bool? IsExpanded { get; set; }
        [JsonProperty(PropertyName = "iconAlignment", NullValueHandling = NullValueHandling.Ignore)]
        public string IconAlignment { get; set; }
        [JsonProperty(PropertyName = "showDividerLine", NullValueHandling = NullValueHandling.Ignore)]
        public bool ShowDividerLine { get; set; }
    }
#endif
}
