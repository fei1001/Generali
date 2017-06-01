// <copyright file="Versions_Cloning.cs" company="">
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author></author>
// <date>5/26/2017 9:42:02 AM</date>
// <summary>Implements the Versions_Cloning Workflow Activity.</summary>
namespace Generali.Workflows
{
    using System;
    using System.Activities;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Workflow;
    using Microsoft.Xrm.Sdk.Query;
    using System.Collections.Generic;
    using Generali.Workflows.BusinessModel;
    public sealed class Versions_Cloning : CodeActivity
    {
        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered Versions_Cloning.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("Versions_Cloning.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);


            try
            {
                Guid tempId = new Guid();
                Guid _newTargetId;
                //List<ChildEntity> ergo_CommissionStatement = new List<ChildEntity>() { new ChildEntity{EntityName = "ergo_commissionstatement", EntityPrimaryKey = "ergo_commissionstatementid" } };
               
                
                //List<string> childEntity = new List<string>() { };
               // List<string> childEntityPrimaryKey = new List<string>(){ };
                // TODO: Implement your custom Workflow business logic.
                Entity _TargetEnttiy;
                _TargetEnttiy = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
                if (_TargetEnttiy == null)
                    return;
                tracingService.Trace("Target Entity Found: " + context.PrimaryEntityName + " ID: " + context.PrimaryEntityId);
                Guid _targetId = _TargetEnttiy.GetAttributeValue<Guid>("ergo_versionid");
                var _targetCurrentId = _TargetEnttiy.GetAttributeValue<int>("ergo_currentversionid");
                var _newTargetCurrentId = _targetCurrentId + 1;
                _TargetEnttiy.Attributes.Remove("ergo_versionid");
                _TargetEnttiy.Attributes.Remove("statuscode");
                _TargetEnttiy.Attributes.Remove("statecode");
                _TargetEnttiy.Id = tempId;
                _TargetEnttiy.Attributes["ergo_currentversionid"] = _newTargetCurrentId;
                _TargetEnttiy.Attributes["ergo_previousrenewalvid"] = _targetCurrentId;
                _newTargetId = service.Create(_TargetEnttiy);
                Entity _team = Helper.RetrieveCRMRecord(service, "team", "name", "Ops",new List<string>(){"name"} );
                tracingService.Trace("team");

                //==================

                CloneChildRecords(service, "ergo_commissionstatement", "ergo_versionid", _targetId, _newTargetId, "ergo_commissionstatementid", "ergo_version");
                CloneChildRecords(service, "ergo_iptstatement", "ergo_versionid", _targetId, _newTargetId, "ergo_iptstatementid", "ergo_version");
                CloneChildRecords(service, "ergo_generaladjustment", "ergo_versionid", _targetId, _newTargetId, "ergo_generaladjustmentid", "ergo_version");
                CloneChildRecords(service, "ergo_policy", "ergo_versionid", _targetId, _newTargetId, "ergo_policyid", "ergo_version");
                CloneChildRecords(service, "ergo_underwritingdecision", "ergo_versionid", _targetId, _newTargetId, "ergo_underwritingdecisionid", "ergo_version");



                ///============


          

            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw new InvalidWorkflowException("There was error. Message:" + e.Message);
            }

            tracingService.Trace("Exiting Versions_Cloning.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        private void CloneChildRecords(IOrganizationService service, string childEntityName,string parentLookupKey,Guid parentId, Guid newParentId, 
            string childEntityPrimaykey, string parentEntityName) {

            Entity _team = Helper.RetrieveCRMRecord(service, "team", "name", "Ops", new List<string>() { "name" });
            //tracingService.Trace("team");
            DataCollection<Entity> ergo_CommissionStatement = Helper.RetrieveCRMRecordsAllCols(service, childEntityName, parentLookupKey, parentId.ToString());
            if (ergo_CommissionStatement != null)
            {
                foreach (var item in ergo_CommissionStatement)
                {
                    Guid tempItemId = new Guid();
                    item.Attributes.Remove(parentLookupKey);
                    item.Attributes.Remove(childEntityPrimaykey);
                    item.Attributes.Remove("statecode");
                    item.Attributes.Remove("statuscode");
                    item.Id = tempItemId;
                    item.Attributes[parentLookupKey] = new EntityReference(parentEntityName, newParentId);
                    Guid newItemId = service.Create(item);

                    if (_team != null)
                    {
                        item.Attributes["Ownerid"] = new EntityReference("team", _team.Id);
                    }

                }

            }
        }
    }
}