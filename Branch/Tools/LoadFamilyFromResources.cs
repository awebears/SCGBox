using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Branch.Tools
{
    public static class LoadFamilyFromResources
    {
        public static void WriteToDisk(string path, Stream stream)
        {
            byte[] bs = new byte[stream.Length];
            stream.Read(bs, 0, bs.Length);
            File.WriteAllBytes(path, bs);
        }
        public static void DeleteFormDisk(string path)
        {
            File.Delete(path);
        }
    }
}
