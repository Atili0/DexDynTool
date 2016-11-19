using System.Collections;
using System.Collections.Generic;

namespace DeXrm.Console
{
    using System;
    using System.Windows.Forms;
    using Microsoft.Xrm.Sdk.Metadata.Query;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Tooling.Connector;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Metadata;

    public class DynamicsCRM
    {
        public void CreateEntity(CrmServiceClient services, DxEntity entity)
        {
            try
            {
                CreateEntityRequest createrequest = new CreateEntityRequest
                {
                    Entity = new EntityMetadata
                    {
                        SchemaName = entity.PSchema,
                        DisplayName = new Microsoft.Xrm.Sdk.Label(entity.PName, 3082),
                        DisplayCollectionName = new Microsoft.Xrm.Sdk.Label(entity.PNamePlural, 3082),
                        Description = new Microsoft.Xrm.Sdk.Label(entity.PDescription, 3082),
                        OwnershipType = entity.POwnershipType,
                        IsActivity = entity.IsActivity,
                        IsConnectionsEnabled = entity.IsConnectionsEnabled,
                        // IsActivityParty = true envio de correo
                    },

                    PrimaryAttribute = new StringAttributeMetadata
                    {
                        SchemaName = entity.DxAttribute.PSchemaName,
                        RequiredLevel = entity.DxAttribute.PRequiredLevel,
                        MaxLength = 100,
                        FormatName = entity.DxAttribute.PFormatName,
                        DisplayName = new Microsoft.Xrm.Sdk.Label(entity.DxAttribute.PDisplayName, 3082),
                        Description = new Microsoft.Xrm.Sdk.Label(entity.PDescription, 3082)
                    }

                };
                OrganizationResponse organizationResponse = services.ExecuteCrmOrganizationRequest(createrequest);
                new Util().CreateLog(string.Format("Entity created at {0} {1} with id {2}",
                    DateTime.Now.ToShortDateString(), DateTime.Now.ToShortTimeString(),
                    ((CreateEntityResponse) organizationResponse).EntityId));
            }
            catch (System.Exception exception)
            {
                var error = exception.Message;
                throw;
            }
        }

        public void GetEntityList(CrmServiceClient service, string logicalName)
        {
            RetrieveEntityRequest retrieveRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Entity,
                LogicalName = logicalName
            };
            RetrieveEntityResponse retrieveEntityResponse = (RetrieveEntityResponse)service.OrganizationServiceProxy.Execute(retrieveRequest);
        }

        public void CreateAttribute(CrmServiceClient service, string entityname, List<DxEntityAttribute> entityAttributes )
        {
            #region Attribute Metadata
            StringAttributeMetadata dxStringAttributeMetadata = null;
            DecimalAttributeMetadata dxDecimalAttributeMetadata = null;
            IntegerAttributeMetadata dxIntegerAttributeMetadata = null;
            MoneyAttributeMetadata dxMoneyAttributeMetadata = null;
            DateTimeAttributeMetadata dxDateTimeAttributeMetadata = null;
            #endregion
            AttributeMetadata attributeMetadata = null;

            foreach (DxEntityAttribute dxEntityAttribute in entityAttributes)
            {
                switch (dxEntityAttribute.PType)
                {
                    case 1:
                        dxStringAttributeMetadata = new StringAttributeMetadata
                        {
                            SchemaName = String.Format("{0}_{1}", entityname.Substring(0, entityname.IndexOf("_")), 
                                dxEntityAttribute.PSchemaName.ToLower().Trim()), 
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                            MaxLength = dxEntityAttribute.StringAttribute.PMaxLength,
                            FormatName = dxEntityAttribute.StringAttribute.PStringFormatName,
                            DisplayName =
                                new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDisplayName, 3082),
                            Description =
                                new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDescription, 3082),
                        };
                        attributeMetadata = dxStringAttributeMetadata;
                        break;
                    case 2:
                        dxDecimalAttributeMetadata = new DecimalAttributeMetadata
                        {
                            SchemaName = String.Format("{0}_{1}", entityname.Substring(0, entityname.IndexOf("_")),
                                dxEntityAttribute.PSchemaName.ToLower().Trim()),
                            DisplayName = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDisplayName, 3082),
                            Description = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDescription, 3082),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),
                           
                            MaxValue = dxEntityAttribute.DecimalAttribute.PMaxValue,
                            MinValue = dxEntityAttribute.DecimalAttribute.PMinValue,
                            Precision = dxEntityAttribute.DecimalAttribute.PPrecision

                        };
                        attributeMetadata = dxDecimalAttributeMetadata;
                        break;
                    case 3:
                        dxIntegerAttributeMetadata = new IntegerAttributeMetadata
                        {
                            SchemaName = String.Format("{0}_{1}", entityname.Substring(0, entityname.IndexOf("_")),
                                dxEntityAttribute.PSchemaName.ToLower().Trim()),
                            DisplayName = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDisplayName, 3082),
                            Description = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDescription, 3082),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                            MaxValue = dxEntityAttribute.IntegerAttribute.PMaxValue,
                            MinValue = dxEntityAttribute.IntegerAttribute.PMinValue,
                            Format = dxEntityAttribute.IntegerAttribute.PFormat

                        };
                        attributeMetadata = dxIntegerAttributeMetadata;
                        break;
                    case 4:
                        dxMoneyAttributeMetadata = new MoneyAttributeMetadata
                        {
                            SchemaName = String.Format("{0}_{1}", entityname.Substring(0, entityname.IndexOf("_")),
                                dxEntityAttribute.PSchemaName.ToLower().Trim()),
                            DisplayName = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDisplayName, 3082),
                            Description = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDescription, 3082),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                            MaxValue = dxEntityAttribute.MoneyAttribute.PMaxValue,
                            MinValue = dxEntityAttribute.MoneyAttribute.PMinValue,
                            Precision = dxEntityAttribute.MoneyAttribute.PPrecision,
                            PrecisionSource = dxEntityAttribute.MoneyAttribute.PPrecisionSource,
                            ImeMode = dxEntityAttribute.MoneyAttribute.PImeMode
                        };

                        attributeMetadata = dxMoneyAttributeMetadata;
                        break;
                    case 5:
                        dxDateTimeAttributeMetadata = new DateTimeAttributeMetadata
                        {
                            SchemaName = String.Format("{0}_{1}", entityname.Substring(0, entityname.IndexOf("_")),
                                dxEntityAttribute.PSchemaName.ToLower().Trim()),
                            DisplayName = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDisplayName, 3082),
                            Description = new Microsoft.Xrm.Sdk.Label(dxEntityAttribute.PDescription, 3082),
                            RequiredLevel = new AttributeRequiredLevelManagedProperty(AttributeRequiredLevel.None),

                            Format = dxEntityAttribute.DateTimeAttribute.PFormat,
                            ImeMode = dxEntityAttribute.DateTimeAttribute.PImeMode

                        };
                        attributeMetadata = dxDateTimeAttributeMetadata;
                        break;

                }

                CreateAttributeRequest createBankNameAttributeRequest = new CreateAttributeRequest
                {
                    EntityName = entityname,
                    Attribute = attributeMetadata,
                };

                service.OrganizationServiceProxy.Execute(createBankNameAttributeRequest);
            }
        }
    }
}
