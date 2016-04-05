using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    public class ProvisioningEventArgs:
        EventArgs
    {
        public BaseModel Model { get; private set; }
        public ClientObject ClientObject { get; private set; }

        public bool Cancel { get; set; }

        public ProvisioningEventArgs(BaseModel model, ClientObject clientObj)
        {
            this.Cancel = false;
            this.Model = model;
            this.ClientObject = clientObj;
        }
    }

    public class ProvisioningEventExecutor
    {
        private enum ProvisionEventType
        {
            PreEvent,
            PostEvent
        }        

        private class ProvisioningEventEntry
        {
            public ProvisionEventType EventType { get; set; }
            public Handlers HandlerType { get; set; }
            public Action<ProvisioningEventArgs> Action { get; set; }

            public bool Execute(ProvisioningTemplate template, BaseModel model, ClientObject clientObject)
            {
                ProvisioningEventArgs eventArgs = new ProvisioningEventArgs(model, clientObject);
                this.Action(eventArgs);

                return !eventArgs.Cancel;
            }
        }

        private List<ProvisioningEventEntry> m_events = null;

        public void RegisterPreProvisionEvent(Handlers handler, Action<ProvisioningEventArgs> fn)
        {
            RegisterEvent(handler, fn, ProvisionEventType.PreEvent);
        }

        public void RegisterPostProvisionEvent(Handlers handler, Action<ProvisioningEventArgs> fn)
        {
            RegisterEvent(handler, fn, ProvisionEventType.PostEvent);
        }

        public bool ExecutePreProvisionEvent(Handlers handler, ProvisioningTemplate template, BaseModel model, ClientObject clientObject)
        {
            bool result = ExecuteEvent(handler, template, model, clientObject, ProvisionEventType.PreEvent);
            return result;
        }

        public void ExecutePostProvisionEvent(Handlers handler, ProvisioningTemplate template, BaseModel model, ClientObject clientObject)
        {
            ExecuteEvent(handler, template, model, clientObject, ProvisionEventType.PostEvent);
        }

        private void RegisterEvent(Handlers handler, Action<ProvisioningEventArgs> fn, ProvisionEventType eventType)
        {
            if (null == m_events)
            {
                m_events = new List<ProvisioningEventEntry>();
            }
            m_events.Add(new ProvisioningEventEntry()
            {
                HandlerType = handler,
                Action = fn,
                EventType = eventType
            });
        }

        private bool ExecuteEvent( Handlers handler, ProvisioningTemplate template, BaseModel model, ClientObject clientObject, ProvisionEventType eventType)
        {
            bool success = true;
            if (null != m_events)
            {
                for (var i = 0; i < m_events.Count; ++i)
                {
                    ProvisioningEventEntry eventEntry = m_events[i];
                    if ((eventEntry.EventType == eventType) &&
                        (eventEntry.HandlerType.HasFlag( handler )))
                    {
                        if (!eventEntry.Execute(template, model, clientObject))
                        {
                            success = false;
                        }
                    }
                }
            }
            return success;
        }
    }
}
