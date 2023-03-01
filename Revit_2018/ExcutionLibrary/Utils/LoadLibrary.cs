using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;

namespace Revit_2018.ExcutionLibrary.Utils
{
    #region LoadAssembly
    internal class LoadLibrary
    {
        private string assemblyName;//在解决方案中的全路径

        public LoadLibrary(string assemblyName)
        {
            this.assemblyName = assemblyName;
        }
        internal void Subscribe()
        {
            AppDomain.CurrentDomain.AssemblyResolve += LoadEmbedAssembly;
        }
        internal void Unsubscribe()
        {
            AppDomain.CurrentDomain.AssemblyResolve -= LoadEmbedAssembly;
        }
        private Assembly LoadEmbedAssembly(object sender, ResolveEventArgs args)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(assemblyName);

            byte[] bs = new byte[stream.Length];
            stream.Read(bs, 0, bs.Length);
            return Assembly.Load(bs);
        }
    }
    #endregion
}
