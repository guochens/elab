using System;

namespace WindowsFormsApplication1
{
    internal class ManagementClass
    {
        private string v;

        public ManagementClass(string v)
        {
            this.v = v;
        }

        internal ManagementObjectCollection GetInstances()
        {
            throw new NotImplementedException();
        }
    }
}