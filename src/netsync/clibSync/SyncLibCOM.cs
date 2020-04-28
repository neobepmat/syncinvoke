
using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;

namespace clibSync.com
{
    [ClassInterface(ClassInterfaceType.None)]
    [ComSourceInterfaces(typeof(ISyncLibComEvents))]
    public class SyncLibCOM : ISyncLibCom
    {
        public delegate void EventHandler(string msg);
        public event EventHandler OnSyncLibEvent;

        public event EventHandler OnBeforeEvent;
        public event EventHandler OnAfterEvent;

        public void Init()
        {
            Timer loopEvent = new Timer(callbackLoopEvent, null, new TimeSpan(0,0,5), new TimeSpan(0,0,5));
        }

        private void callbackLoopEvent(object state)
        {
            FireEvent("Message sent at: " + DateTime.Now.ToString());
        }

        public void Method()
        {
            throw new System.NotImplementedException();
        }

        public void FireEvent(string msg)
        {
            OnBeforeEvent?.Invoke("Before Event is: " + DateTime.Now.ToString());
            System.Diagnostics.Debug.WriteLine("Before Event is: " + DateTime.Now.ToString());
            Task.Run(() => OnSyncLibEvent?.Invoke(msg));
            OnAfterEvent?.Invoke("After Event is: " + DateTime.Now.ToString());
            System.Diagnostics.Debug.WriteLine("After Event is: " + DateTime.Now.ToString());
        }
    }

    [Guid("6FF0D360-43DC-4701-A9F4-7BE912D02A9F")]
    [ComVisible(true)]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface ISyncLibComEvents
    {
        [DispId(10)]
        void OnSyncLibEvent(string msg);
    }

    [Guid("CDCD063F-3719-42FA-BEDB-04A0768E350F")]
    [ComVisible(true)]
    public interface ISyncLibCom
    {
        [DispId(10)]
        void Init();
    }
}
