using System;

namespace DeXrm.Console
{
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Metadata;
    using Microsoft.Xrm.Tooling.Connector;

    public class DxEntity
    {
        public string PName { get; set; }
        public string PSchema { get; set; }
        public string PNamePlural { get; set; }
        public string PDescription { get; set; }
        public bool IsActivity { get; set; }
        public OwnershipTypes POwnershipType { get; set; }
        public BooleanManagedProperty IsConnectionsEnabled { get; set; }
        public DxAttribute DxAttribute { get; set; }
        public string PSolutionSchema { get; set; }
    }

    public class DxAttribute
    {
        public string PSchemaName { get; set; }
        public string PDisplayName { get; set; }
        public StringFormatName PFormatName { get; set; }
        public AttributeRequiredLevelManagedProperty PRequiredLevel { get; set; }
        

    }

    public class DxEntityAttribute
    {
        public string PSchemaName { get; set; }
        public string PDisplayName { get; set; }
        public int PType { get; set; }
        public AttributeRequiredLevelManagedProperty PAttributeRequiredLevelManagedProperty { get; set; }
        public string PDescription { get; set; }
        public DxStringAttribute StringAttribute { get; set; }
        public DxDecimalAttribute DecimalAttribute { get; set; }
        public DxIntegerAttributeMetadata IntegerAttribute { get; set; }
        public DxMoneyAttributeMetadata MoneyAttribute { get; set; }
        public DxDateTimeAttributeMetadata DateTimeAttribute { get; set; }
    }

    public class DxStringAttribute
    {
        public int PMaxLength { get; set; }
        public StringFormatName PStringFormatName { get; set; } 
    }

    public class DxDecimalAttribute
    {
        public int PMaxValue { get; set; }
        public int PMinValue { get; set; }
        public int PPrecision { get; set; }
    }

    public class DxIntegerAttributeMetadata
    {
        public int PMaxValue { get; set; }
        public int PMinValue { get; set; }
        public IntegerFormat PFormat { get; set; }
    }

    public class DxMoneyAttributeMetadata
    {
        public Double PMaxValue { get; set; }
        public Double PMinValue { get; set; }
        public int PPrecision { get; set; }
        public int PPrecisionSource { get; set; }
        public ImeMode PImeMode { get; set; }
    }

    public class DxDateTimeAttributeMetadata
    {
        public DateTimeFormat PFormat { get; set; }
        public ImeMode PImeMode { get; set; }
    }

    public class Connection
    {
        public static CrmServiceClient CrmSvc { get; set; }
    }


}
