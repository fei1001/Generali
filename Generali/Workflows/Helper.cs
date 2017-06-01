using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;

namespace Generali.Workflows
{
   public static class Helper
    {
        /// <summary>RetrieveCRMRecord
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="EntityName"></param>
        /// <param name="SearchAttribute"></param>
        /// <param name="SearchValue"></param>
        /// <param name="AttributesToRetrieve"></param>
        /// <returns></returns>
        public static Entity RetrieveCRMRecord(IOrganizationService orgService, string EntityName, string SearchAttribute, string SearchValue, List<string> AttributesToRetrieve)
        {
            ColumnSet culmnS = new ColumnSet();
            foreach (string attribute in AttributesToRetrieve)
                culmnS.Columns.Add(attribute);

            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = SearchAttribute;
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(SearchValue);

            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(ce);

            QueryExpression entQuery = new QueryExpression();
            entQuery.EntityName = EntityName;
            entQuery.ColumnSet = culmnS;
            entQuery.Criteria = filter;

            DataCollection<Entity> retrievedEntities = orgService.RetrieveMultiple(entQuery).Entities;

            if (retrievedEntities.Count > 0)
                return retrievedEntities[0];

            return null;
        }

        /// <summary>RetrieveCRMRecords
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="EntityName"></param>
        /// <param name="SearchAttribute"></param>
        /// <param name="SearchValue"></param>
        /// <param name="AttributesToRetrieve"></param>
        /// <returns></returns>
        public static DataCollection<Entity> RetrieveCRMRecords(IOrganizationService orgService, string EntityName, string SearchAttribute,
                                                                string SearchValue, List<string> AttributesToRetrieve)
        {
            ColumnSet columnS = new ColumnSet();
            foreach (string attribute in AttributesToRetrieve)
                columnS.Columns.Add(attribute);

            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = SearchAttribute;
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(SearchValue);

            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(ce);

            QueryExpression entQuery = new QueryExpression();
            entQuery.EntityName = EntityName;
            entQuery.ColumnSet = columnS;
            entQuery.Criteria = filter;

            DataCollection<Entity> retrievedEntities = orgService.RetrieveMultiple(entQuery).Entities;

            return retrievedEntities;
        }

        /// <summary>RetrieveCRMRecords
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="EntityName"></param>
        /// <param name="SearchAttribute"></param>
        /// <param name="SearchValue"></param>
        /// <param name="AttributesToRetrieve"></param>
        /// <returns></returns>
        public static DataCollection<Entity> RetrieveCRMRecordsAllCols(IOrganizationService orgService, string EntityName, string SearchAttribute,
                                                                string SearchValue)
        {
            ColumnSet columnS = new ColumnSet(true);
            ConditionExpression ce = new ConditionExpression();
            ce.AttributeName = SearchAttribute;
            ce.Operator = ConditionOperator.Equal;
            ce.Values.Add(SearchValue);

            FilterExpression filter = new FilterExpression();
            filter.Conditions.Add(ce);

            QueryExpression entQuery = new QueryExpression();
            entQuery.EntityName = EntityName;
            entQuery.ColumnSet = columnS;
            entQuery.Criteria = filter;

            DataCollection<Entity> retrievedEntities = orgService.RetrieveMultiple(entQuery).Entities;

            return retrievedEntities;
        }


        /// <summary>UpdateRecord
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entityName"></param>
        /// <param name="entityKey"></param>
        /// <param name="attributeList"></param>
        public static void UpdateRecord(IOrganizationService orgService, string entityName, Guid entityKey, List<object[]> attributeList)
        {
            Entity ent = new Entity();
            ent.LogicalName = entityName;

            foreach (object[] attribute in attributeList)
            {
                ent.Attributes[(string)attribute[0]] = attribute[1];
            }

            ent.Attributes[entityName + "id"] = entityKey;
            orgService.Update(ent);
        }

        /// <summary>CreateManyToManyRelation
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="relationshipName"></param>
        /// <param name="entity1"></param>
        /// <param name="id1"></param>
        /// <param name="entity2"></param>
        /// <param name="id2"></param>
        public static void CreateManyToManyRelation(IOrganizationService orgService, string relationshipName, string entity1, Guid id1, string entity2, Guid id2)
        {
            EntityReferenceCollection relatedEntities = new EntityReferenceCollection();
            relatedEntities.Add(new EntityReference(entity1, id1));

            Relationship relationship = new Relationship(relationshipName);

            orgService.Associate(entity2, id2, relationship, relatedEntities);

        }

        /// <summary>GetOpportunityIdInOPPClose
        /// 
        /// </summary>
        /// <param name="context"></param>
        /// <returns></returns>
        public static Guid GetOpportunityIdInOPPClose(IPluginExecutionContext context)
        {
            Guid opportunityId = Guid.Empty;
            if (context.InputParameters.Contains("OpportunityClose") &&
                context.InputParameters["OpportunityClose"] is Entity)
            {
                Entity entity = (Entity)context.InputParameters["OpportunityClose"];
                if (entity.Attributes.Contains("opportunityid") &&
                   entity.Attributes["opportunityid"] != null)
                {
                    EntityReference entityRef = (EntityReference)entity.Attributes["opportunityid"];
                    if (entityRef.LogicalName == "opportunity")
                    {
                        opportunityId = entityRef.Id;
                    }
                }
            }
            return opportunityId;
        }

        /// <summary>AssignEntity
        /// 
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="record"></param>
        /// <param name="asignee"></param>
        public static void AssignEntity(IOrganizationService orgService, EntityReference record, EntityReference asignee)
        {
            AssignRequest assignReq = new AssignRequest();
            assignReq.Assignee = asignee;
            assignReq.Target = record;
            orgService.Execute(assignReq);
        }

        /// <summary>Deactivate a record
        /// 
        /// </summary>
        /// <param name="organizationService"></param>
        /// <param name="entityName"></param>
        /// <param name="recordId"></param>
        public static void DeactivateRecord(IOrganizationService organizationService, string entityName, Guid recordId)
        {
            var cols = new ColumnSet(new[] { "statecode", "statuscode" });

            //Check if it is Active or not
            var entity = organizationService.Retrieve(entityName, recordId, cols);

            if (entity != null && entity.GetAttributeValue<OptionSetValue>("statecode").Value == 0)
            {
                //StateCode = 1 and StatusCode = 2 for deactivating Account or Contact
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = recordId,
                        LogicalName = entityName,
                    },
                    State = new OptionSetValue(1),
                    Status = new OptionSetValue(2)
                };
                organizationService.Execute(setStateRequest);
            }
        }

        /// <summary>Activate a record
        /// 
        /// </summary>
        /// <param name="organizationService"></param>
        /// <param name="entityName"></param>
        /// <param name="recordId"></param>
        public static void ActivateRecord(IOrganizationService organizationService, string entityName, Guid recordId)
        {
            var cols = new ColumnSet(new[] { "statecode", "statuscode" });

            //Check if it is Inactive or not
            var entity = organizationService.Retrieve(entityName, recordId, cols);

            if (entity != null && entity.GetAttributeValue<OptionSetValue>("statecode").Value == 1)
            {
                //StateCode = 0 and StatusCode = 1 for activating Account or Contact
                SetStateRequest setStateRequest = new SetStateRequest()
                {
                    EntityMoniker = new EntityReference
                    {
                        Id = recordId,
                        LogicalName = entityName,
                    },
                    State = new OptionSetValue(0),
                    Status = new OptionSetValue(1)
                };
                organizationService.Execute(setStateRequest);
            }
        }

        /// <summary>RetrieveOptionSetTextByValue
        /// fieldValue = ((OptionSetValue)entity.Attributes[field]).Value.ToString();
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entityName"></param>
        /// <param name="fieldName"></param>
        /// <param name="_RetrieveAsIfPublished"></param>
        /// <param name="fieldValue"></param>
        /// <returns></returns>
        public static String RetrieveOptionSetTextByValue(IOrganizationService orgService, string entityName, string fieldName, bool _RetrieveAsIfPublished, int fieldValue)
        {
            string _OptionSetText = null;
            RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
            {
                EntityLogicalName = entityName,
                LogicalName = fieldName,
                RetrieveAsIfPublished = _RetrieveAsIfPublished

            };

            RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)orgService.Execute(attributeRequest);
            EnumAttributeMetadata attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            foreach (OptionMetadata om in attributeMetadata.OptionSet.Options)
            {
                if (om.Value == fieldValue)
                {
                    _OptionSetText = om.Label.UserLocalizedLabel.Label;
                }
            }

            return _OptionSetText;
        }

        /// <summary>SetOptionSetByValue
        ///int value = ((OptionSetValue)entity["yourattributename"]).Value;
        ///String text = entity.FormattedValues["yourattributename"].ToString();
        /// </summary>
        /// <param name="orgService"></param>
        /// <param name="entity"></param>
        /// <param name="fieldName"></param>
        /// <param name="fieldValue"></param>
        public static void SetOptionSetByValue(IOrganizationService orgService, Entity entity, string fieldName, int fieldValue)
        {

            RetrieveOptionSetRequest Request = new RetrieveOptionSetRequest
            {
                Name = fieldName
            };

            RetrieveOptionSetResponse Response = (RetrieveOptionSetResponse)orgService.Execute(Request);
            OptionSetMetadata md = (OptionSetMetadata)Response.OptionSetMetadata;
            var currentOptions = md.Options.ToArray();
            foreach (var item in currentOptions)
            {
                if (item.Value == fieldValue)
                {
                    OptionSetValue val = new OptionSetValue((int)item.Value);
                    var _entity = new Entity(entity.LogicalName) { Id = entity.Id };
                    _entity[fieldName] = val;
                    orgService.Update(_entity);
                }
            }




            //RetrieveAttributeRequest attributeRequest = new RetrieveAttributeRequest
            //{
            //    EntityLogicalName = entity.LogicalName,
            //    LogicalName = fieldName,
            //    RetrieveAsIfPublished = _RetrieveAsIfPublished

            //};

            //RetrieveAttributeResponse attributeResponse = (RetrieveAttributeResponse)orgService.Execute(attributeRequest);
            //EnumAttributeMetadata attributeMetadata = (EnumAttributeMetadata)attributeResponse.AttributeMetadata;

            //foreach (OptionMetadata om in attributeMetadata.OptionSet.Options)
            //{
            //    if (om.Value == fieldValue)
            //    {
            //        OptionSetValue op = new OptionSetValue(om.Value.Value);
            //        entity.Attributes[fieldName] = op;
            //        orgService.Update(entity);
            //    }
            //}
        }

        ///

        public static string GetAttributeString(this Entity entity, string attributeName)
        {
            string returnValue = string.Empty;
            if (entity.Contains(attributeName) && entity[attributeName] != null)
            {

                string type = entity[attributeName].GetType().Name.ToLower();
                switch (type)
                {
                    case "string":
                        returnValue = entity.GetAttributeValue<string>(attributeName);
                        break;
                    case "datetime":
                        returnValue = entity.GetAttributeValue<DateTime>(attributeName).ToShortDateString();
                        break;
                    case "decimal":
                        returnValue = entity.GetAttributeValue<decimal>(attributeName).ToString();
                        break;
                    case "entityreference":
                        returnValue = entity.GetAttributeValue<EntityReference>(attributeName).Name;
                        break;
                    case "int32":
                        returnValue = entity.GetAttributeValue<int>(attributeName).ToString();
                        break;
                }
            }
            return returnValue;
        }

        public static T GetAttributeValue<T>(this IPluginExecutionContext context, string attributeName)
        {
            T value = default(T);
            if (context.InputParameters != null && context.InputParameters.Contains("Target") &&
                 context.InputParameters["Target"] is Entity)
            {
                Entity e = context.InputParameters["Target"] as Entity;
                value = e.GetAttributeValue<T>(attributeName);
            }

            if (value == null && context.PreEntityImages != null && context.PreEntityImages.Contains("img") &&
                context.PreEntityImages["img"].GetAttributeValue<T>(attributeName) != null)
                value = context.PreEntityImages["img"].GetAttributeValue<T>(attributeName);

            if (value == null && context.PostEntityImages != null && context.PostEntityImages.Contains("img") &&
                context.PostEntityImages["img"].GetAttributeValue<T>(attributeName) != null)
                value = context.PostEntityImages["img"].GetAttributeValue<T>(attributeName);
            return value;
        }

        public static T GetAliasedValue<T>(this Entity entity, string attributeName)
        {
            AliasedValue value = entity.GetAttributeValue<AliasedValue>(attributeName);
            if (value == null)
                return default(T);
            return (T)value.Value;
        }

        public static T GetAliasedValueWithDefault<T>(this Entity entity, string attributeName)
        {
            AliasedValue value = entity.GetAttributeValue<AliasedValue>(attributeName);
            if (value == null)
                return default(T);
            if (value.Value == null)
                return default(T);
            return (T)value.Value;
        }

        public static string GetDisplayName(this AttributeMetadata[] attributes, string fieldName)
        {
            string displayName = string.Empty;

            var attribute = attributes.Where(x => x.SchemaName.ToLower() == fieldName.ToLower()).FirstOrDefault();
            if (attribute != null)
            {
                LocalizedLabel label = attribute.DisplayName.LocalizedLabels.FirstOrDefault();
                if (label != null) displayName = label.Label;
            }
            return displayName;
        }

    }
}
