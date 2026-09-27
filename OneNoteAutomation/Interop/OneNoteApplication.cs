using System;
using System.Runtime.InteropServices;

namespace OneNoteAutomation.Interop
{
    public sealed partial class OneNoteApplication
    {
        private object instance;
        private IOneNoteApplication application;
        private readonly Action<object> release;

        public OneNoteApplication()
            : this(Activator.CreateInstance(Type.GetTypeFromProgID("OneNote.Application", true)),
                ownedInstance => Marshal.ReleaseComObject(ownedInstance))
        {
        }

        internal OneNoteApplication(object instance, Action<object> release)
        {
            if (instance == null) throw new ArgumentNullException("instance");
            if (release == null) throw new ArgumentNullException("release");
            this.instance = instance;
            this.release = release;
            try
            {
                application = (IOneNoteApplication)instance;
            }
            catch
            {
                release(instance);
                this.instance = null;
                throw;
            }
        }

        private IOneNoteApplication Application
        {
            get
            {
                if (application == null) throw new ObjectDisposedException("OneNoteApplication");
                return application;
            }
        }

        public void Dispose()
        {
            if (instance == null) return;
            object ownedInstance = instance;
            application = null;
            instance = null;
            release(ownedInstance);
        }
    }
}
