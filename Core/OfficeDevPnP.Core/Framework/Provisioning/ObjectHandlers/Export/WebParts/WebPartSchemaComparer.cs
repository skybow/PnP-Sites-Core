using System;
using System.Xml.Linq;
using OfficeDevPnP.Core.Framework.Provisioning.Model;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.Export.WebParts
{
    public interface IWebPartSchemaComparer
    {
        bool IsDefaultWebPart(Model.WebPart webPart);
    }

    public abstract class WebPartSchemaComparer:
        IWebPartSchemaComparer
    {
        #region Properties

        protected abstract string WebPartDefaultSchema { get; }

        protected abstract WebPartPropertiesCleanerBase WebPartPropertiesCleaner { get; }

        protected virtual bool IgnoreAttributes
        {
            get
            {
                return true;
            }
        }

        #endregion //Properties

        #region Methods

        public virtual bool IsDefaultWebPart(WebPart webPart)
        {
            XElement xmlDefWebPart = XElement.Parse(this.WebPartDefaultSchema);
            XElement xmlWebPart = XElement.Parse(webPart.Contents);

            PrepareToCompare(xmlDefWebPart);
            PrepareToCompare(xmlWebPart);

            bool result = XElement.DeepEquals(xmlDefWebPart, xmlWebPart);
            return result;
        }

        protected virtual void PrepareToCompare(XElement xmlWebPart)
        {
            if (this.IgnoreAttributes)
            {
                xmlWebPart.Attributes().Remove();
            }

            this.WebPartPropertiesCleaner.CleanDefaultProperties(xmlWebPart);
        }

        public static IWebPartSchemaComparer CreateTypedComparer(WebPart webpart)
        {
            IWebPartSchemaComparer result = null;

            if (WebPartsModelProvider.IsV3FormatXml(webpart.Contents))
            {
                if (-1 != webpart.Contents.IndexOf(@"<type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart", StringComparison.OrdinalIgnoreCase))
                {
                    result = new XsltListViewWebPartComparer();
                }
            }
            else
            {
                if (-1 != webpart.Contents.IndexOf("<TypeName>Microsoft.SharePoint.WebPartPages.ListFormWebPart</TypeName>", StringComparison.OrdinalIgnoreCase))
                {
                    result = new ListFormWebPartComparer();
                }
            }

            return result;
        }

        #endregion //Methods
    }

    public class XsltListViewWebPartComparer:
        WebPartSchemaComparer
    {
        #region Constants

        private const string SCHEMA_DEFAULT = @"
<webParts webpartid=""[WEBPARTID]"">
    <webPart xmlns=""http://schemas.microsoft.com/WebPart/v3"">
      <metaData>
        <type name=""Microsoft.SharePoint.WebPartPages.XsltListViewWebPart, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"" />
        <importErrorMessage>Cannot import this Web Part.</importErrorMessage>
      </metaData>
      <data>
        <properties>
          <property name=""ShowWithSampleData"" type=""bool"">False</property>
          <property name=""Default"" type=""string""></property>
          <property name=""NoDefaultStyle"" type=""string""></property>
          <property name=""CacheXslStorage"" type=""bool"">True</property>
          <property name=""ViewContentTypeId"" type=""string""></property>
          <property name=""XmlDefinitionLink"" type=""string""></property>
          <property name=""ManualRefresh"" type=""bool"">False</property>
          <property name=""ListUrl"" type=""string"" null=""true""></property>
          <property name=""ListId"" type=""System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"">{listid:Calendar}</property>
          <property name=""TitleUrl"" type=""string"">~site/{listurl:Calendar}</property>
          <property name=""EnableOriginalValue"" type=""bool"">False</property>
          <property name=""Direction"" type=""direction"">NotSet</property>
          <property name=""ServerRender"" type=""bool"">False</property>
          <property name=""ViewFlags"" type=""Microsoft.SharePoint.SPViewFlags, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"">Html, TabularView, Mobile</property>
          <property name=""AllowConnect"" type=""bool"">True</property>
          <property name=""ListName"" type=""string"">{{listid:Calendar}}</property>
          <property name=""ListDisplayName"" type=""string"" null=""true""></property>
          <property name=""AllowZoneChange"" type=""bool"">True</property>
          <property name=""ChromeState"" type=""chromestate"">Normal</property>
          <property name=""DisableSaveAsNewViewButton"" type=""bool"">False</property>
          <property name=""ViewFlag"" type=""string""></property>
          <property name=""DataSourceID"" type=""string""></property>
          <property name=""ExportMode"" type=""exportmode"">All</property>
          <property name=""AutoRefresh"" type=""bool"">False</property>
          <property name=""FireInitialRow"" type=""bool"">True</property>
          <property name=""AllowEdit"" type=""bool"">True</property>
          <property name=""Description"" type=""string""></property>
          <property name=""HelpMode"" type=""helpmode"">Modeless</property>
          <property name=""BaseXsltHashKey"" type=""string"" null=""true""></property>
          <property name=""AllowMinimize"" type=""bool"">True</property>
          <property name=""CacheXslTimeOut"" type=""int"">86400</property>
          <property name=""ChromeType"" type=""chrometype"">Default</property>
          <property name=""Xsl"" type=""string"" null=""true""></property>
          <property name=""JSLink"" type=""string"" null=""true""></property>
          <property name=""CatalogIconImageUrl"" type=""string""></property>
          <property name=""SampleData"" type=""string"" null=""true""></property>
          <property name=""UseSQLDataSourcePaging"" type=""bool"">True</property>
          <property name=""TitleIconImageUrl"" type=""string""></property>
          <property name=""PageSize"" type=""int"">-1</property>
          <property name=""ShowTimelineIfAvailable"" type=""bool"">True</property>
          <property name=""Width"" type=""string""></property>
          <property name=""DataFields"" type=""string""></property>
          <property name=""Hidden"" type=""bool"">False</property>
          <property name=""Title"" type=""string""></property>
          <property name=""PageType"" type=""Microsoft.SharePoint.PAGETYPE, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"">PAGE_NORMALVIEW</property>
          <property name=""DataSourcesString"" type=""string""></property>
          <property name=""AllowClose"" type=""bool"">True</property>
          <property name=""InplaceSearchEnabled"" type=""bool"">True</property>
          <property name=""WebId"" type=""System.Guid, mscorlib, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089"">00000000-0000-0000-0000-000000000000</property>
          <property name=""Height"" type=""string""></property>
          <property name=""GhostedXslLink"" type=""string"">main.xsl</property>
          <property name=""DisableViewSelectorMenu"" type=""bool"">False</property>
          <property name=""DisplayName"" type=""string"">All Events</property>
          <property name=""IsClientRender"" type=""bool"">False</property>
          <property name=""XmlDefinition"" type=""string"">&lt;View Name=""{DC250CC6-7129-4DF7-9B63-827D8F00146F}"" MobileView=""TRUE"" Type=""HTML"" DisplayName=""All Events"" Url=""~site/{listurl:Calendar}/AllItems.aspx"" Level=""1"" BaseViewID=""1"" ContentTypeID=""0x"" ImageUrl=""/_layouts/15/images/events.png?rev=23"" &gt;&lt;Query&gt;&lt;OrderBy&gt;&lt;FieldRef Name=""EventDate""/&gt;&lt;/OrderBy&gt;&lt;/Query&gt;&lt;ViewFields&gt;&lt;FieldRef Name=""fRecurrence""/&gt;&lt;FieldRef Name=""WorkspaceLink""/&gt;&lt;FieldRef Name=""LinkTitle""/&gt;&lt;FieldRef Name=""Location""/&gt;&lt;FieldRef Name=""EventDate""/&gt;&lt;FieldRef Name=""EndDate""/&gt;&lt;FieldRef Name=""fAllDayEvent""/&gt;&lt;/ViewFields&gt;&lt;RowLimit Paged=""TRUE""&gt;30&lt;/RowLimit&gt;&lt;JSLink&gt;clienttemplates.js&lt;/JSLink&gt;&lt;XslLink Default=""TRUE""&gt;main.xsl&lt;/XslLink&gt;&lt;Toolbar Type=""Standard""/&gt;&lt;/View&gt;</property>
          <property name=""InitialAsyncDataFetch"" type=""bool"">False</property>
          <property name=""AllowHide"" type=""bool"">True</property>
          <property name=""ParameterBindings"" type=""string"">
            &lt;ParameterBinding Name=""dvt_sortdir"" Location=""Postback;Connection""/&gt;
            &lt;ParameterBinding Name=""dvt_sortfield"" Location=""Postback;Connection""/&gt;
            &lt;ParameterBinding Name=""dvt_startposition"" Location=""Postback"" DefaultValue=""""/&gt;
            &lt;ParameterBinding Name=""dvt_firstrow"" Location=""Postback;Connection""/&gt;
            &lt;ParameterBinding Name=""OpenMenuKeyAccessible"" Location=""Resource(wss,OpenMenuKeyAccessible)"" /&gt;
            &lt;ParameterBinding Name=""open_menu"" Location=""Resource(wss,open_menu)"" /&gt;
            &lt;ParameterBinding Name=""select_deselect_all"" Location=""Resource(wss,select_deselect_all)"" /&gt;
            &lt;ParameterBinding Name=""idPresEnabled"" Location=""Resource(wss,idPresEnabled)"" /&gt;&lt;ParameterBinding Name=""NoAnnouncements"" Location=""Resource(wss,noXinviewofY_LIST)"" /&gt;&lt;ParameterBinding Name=""NoAnnouncementsHowTo"" Location=""Resource(wss,noXinviewofY_DEFAULT)"" /&gt;
          </property>
          <property name=""DataSourceMode"" type=""Microsoft.SharePoint.WebControls.SPDataSourceMode, Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c"">List</property>
          <property name=""AutoRefreshInterval"" type=""int"">60</property>
          <property name=""AsyncRefresh"" type=""bool"">False</property>
          <property name=""HelpUrl"" type=""string""></property>
          <property name=""MissingAssembly"" type=""string"">Cannot import this Web Part.</property>
          <property name=""XslLink"" type=""string"" null=""true""></property>
          <property name=""SelectParameters"" type=""string""></property>
        </properties>
      </data>
    </webPart>
</webParts>";

        private static readonly string[] Ignore_Properties = new string[]
        {
            "ListId",
            "TitleUrl",
            "ListName",
            "DisplayName",
            "ParameterBindings",
            "XmlDefinition",
            "ViewFlags"
        };

        #endregion //Constants

        #region Overrides

        protected override string WebPartDefaultSchema
        {
	        get { return SCHEMA_DEFAULT; }
        }

        protected override WebPartPropertiesCleanerBase WebPartPropertiesCleaner
        {
	        get
            {
                return new V3.V3WebPartPropertiesCleaner(Ignore_Properties);
            }
        }

        #endregion //Overrides
    }

    public class ListFormWebPartComparer:
        WebPartSchemaComparer
    {
        #region Constants

        private const string SCHEMA_DEFAULT = @"
<WebPart 
    xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" 
    xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" 
    xmlns=""http://schemas.microsoft.com/WebPart/v2"" 
    ID=""[WEBPARTID]"" 
    webpartid=""[WEBPARTID]"">
  <Title></Title>
  <FrameType>Default</FrameType>
  <Description></Description>
  <IsIncluded>true</IsIncluded>
  <ZoneID>Main</ZoneID>
  <PartOrder>1</PartOrder>
  <FrameState>Normal</FrameState>
  <Height></Height>
  <Width></Width>
  <AllowRemove>true</AllowRemove>
  <AllowZoneChange>true</AllowZoneChange>
  <AllowMinimize>true</AllowMinimize>
  <AllowConnect>true</AllowConnect>
  <AllowEdit>true</AllowEdit>
  <AllowHide>true</AllowHide>
  <IsVisible>true</IsVisible>
  <DetailLink></DetailLink>
  <HelpLink></HelpLink>
  <HelpMode>Modeless</HelpMode>
  <Dir>Default</Dir>
  <PartImageSmall></PartImageSmall>
  <MissingAssembly>Cannot import this Web Part.</MissingAssembly>
  <PartImageLarge></PartImageLarge>
  <IsIncludedFilter></IsIncludedFilter>
  <ExportControlledProperties>true</ExportControlledProperties>
  <ConnectionID>00000000-0000-0000-0000-000000000000</ConnectionID>
  <ID>g_a05d3cd4_de19_4dbf_b532_4a4a3486b3ee</ID>
  <Assembly>Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
  <TypeName>Microsoft.SharePoint.WebPartPages.ListFormWebPart</TypeName>
  <ListName xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">{{listid:[CURRENTLIST]}}</ListName>
  <ListId xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">{listid:[CURRENTLIST]}</ListId>
  <PageType xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">PAGE_NEWFORM</PageType>
  <FormType xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">8</FormType>
  <ControlMode xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">New</ControlMode>
  <ViewFlag xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">1048576</ViewFlag>
  <ViewFlags xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">Default</ViewFlags>
  <ListItemId xmlns=""http://schemas.microsoft.com/WebPart/v2/ListForm"">0</ListItemId>
</WebPart>";

        private static readonly string[] Ignore_Properties = new string[]
        {
            "ZoneID",
            "PartOrder",
            "MissingAssembly",
            "ID",
            "Assembly",
            "TypeName",
            "PageType",
            "FormType",
            "ControlMode",
            "ListName",
            "ListId",
            "CSRRenderMode" //Some lists be default has different default value. If ListForm is edited in browser, then new TemplateName, Title attribute is added => this property can be excluded.
        };

        #endregion //Constants

        #region Overrides

        protected override string WebPartDefaultSchema
        {
	        get { return SCHEMA_DEFAULT; }
        }

        protected override WebPartPropertiesCleanerBase WebPartPropertiesCleaner
        {
	        get
            {
                return new V2.V2WebPartPropertiesCleaner(Ignore_Properties);
            }
        }

        #endregion //Overrides
    }
}
