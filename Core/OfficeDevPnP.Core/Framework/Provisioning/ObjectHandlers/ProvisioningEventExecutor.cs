using Microsoft.SharePoint.Client;
using MSClient = Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using Model = OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public class ProvisioningEventArgs<TModel, TClient> :
        EventArgs
        where TModel : BaseModel
        where TClient : ClientObject
    {
        public TModel Model { get; private set; }
        public TClient ClientObject { get; private set; }

        public bool Cancel { get; set; }

        public ProvisioningEventArgs(TModel model, TClient clientObj)
        {
            this.Cancel = false;
            this.Model = model;
            this.ClientObject = clientObj;
        }
    }

    public enum ProvisionEventType
    {
        PreProcessor,
        PostProcessor
    }

    public class ProvisioningEventExecutor
    {
        private interface IProvisioningEventEntry
        {
            ProvisionEventType EventType { get; set; }
            Handlers HandlerType { get; set; }
        }

        private class ProvisioningEventEntry<TModel,TClient>:
            IProvisioningEventEntry
            where TModel: BaseModel
            where TClient: ClientObject
        {
            public ProvisionEventType EventType { get; set; }
            public Handlers HandlerType { get; set; }

            public Action<ProvisioningEventArgs<TModel, TClient>> Action { get; set; }

            public bool Execute(ProvisioningTemplate template, TModel model, TClient clientObject)
            {
                ProvisioningEventArgs<TModel, TClient> eventArgs = new ProvisioningEventArgs<TModel, TClient>(model, clientObject);
                this.Action(eventArgs);

                return !eventArgs.Cancel;
            }
        }

        private List<IProvisioningEventEntry> m_events = null;

        public void AttachListEvent(ProvisionEventType eventType, Action<ProvisioningEventArgs<ListInstance, List>> fn)
        {
            RegisterEvent(Handlers.Lists, fn, eventType);
        }

        public void AttachListContentEvent(ProvisionEventType eventType, Action<ProvisioningEventArgs<ListInstance, List>> fn)
        {
            RegisterEvent(Handlers.ListContents, fn, eventType);
        }

        public void AttachFieldEvent(ProvisionEventType eventType, 
            Action<ProvisioningEventArgs<Model.Field, MSClient.Field>> fn)
        {
            RegisterEvent(Handlers.Fields, fn, eventType);
        }        

        public bool ExecutePreProvisionEvent<TModel, TClient>(Handlers handler, ProvisioningTemplate template, TModel model, TClient clientObject)
            where TModel : BaseModel
            where TClient : ClientObject
        {
            bool result = ExecuteEvent(handler, template, model, clientObject, ProvisionEventType.PreProcessor);
            return result;
        }

        public void ExecutePostProvisionEvent<TModel, TClient>(Handlers handler, ProvisioningTemplate template, TModel model, TClient clientObject)
            where TModel : BaseModel
            where TClient : ClientObject
        {
            ExecuteEvent(handler, template, model, clientObject, ProvisionEventType.PostProcessor);
        }

        private void RegisterEvent<TModel, TClient>(Handlers handler, Action<ProvisioningEventArgs<TModel, TClient>> fn, ProvisionEventType eventType)
            where TModel : BaseModel
            where TClient : ClientObject
        {
            if (null == m_events)
            {
                m_events = new List<IProvisioningEventEntry>();
            }
            m_events.Add(new ProvisioningEventEntry<TModel,TClient>()
            {
                HandlerType = handler,
                EventType = eventType,
                Action = fn       
            });
        }

        private bool ExecuteEvent<TModel, TClient>(Handlers handler, ProvisioningTemplate template, TModel model, TClient clientObject, ProvisionEventType eventType)
            where TModel : BaseModel
            where TClient : ClientObject
        {
            bool success = true;
                        
            if (null != m_events)
            {
                for (var i = 0; i < m_events.Count; ++i)
                {
                    IProvisioningEventEntry eventEntry = m_events[i];
                    if ((eventEntry.EventType == eventType) &&
                        (eventEntry.HandlerType == handler))
                    {
                        ProvisioningEventEntry<TModel, TClient> eventEntryTyped = m_events[i] as ProvisioningEventEntry<TModel, TClient>;
                        if (null != eventEntryTyped)
                        {
                            if (!eventEntryTyped.Execute(template, model, clientObject))
                            {
                                success = false;
                            }
                        }
                    }
                }
            }
            return success;
        }
    }
}
