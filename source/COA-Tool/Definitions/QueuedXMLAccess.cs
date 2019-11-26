using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace CoA_Tool.Definitions
{
    class QueuedXMLAccess : IDisposable
    {
        private static readonly Mutex Access = new Mutex();
        public FileStream fileStream { get; private set; }

        public QueuedXMLAccess (string path, FileMode fileMode, FileAccess fileAccess, FileShare fileShare)
        {
            Access.WaitOne();
        }
    }
}
