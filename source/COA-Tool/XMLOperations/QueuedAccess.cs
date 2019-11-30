using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace CoA_Tool.XMLOperations
{
    class QueuedAccess : IDisposable
    {
        private static readonly Mutex Access = new Mutex();
        public FileStream XMLFileStream { get; private set; }

        public QueuedAccess (string path, FileMode fileMode, FileAccess fileAccess, FileShare fileShare)
        {
            Access.WaitOne();
            XMLFileStream = File.Open(path, fileMode, fileAccess, fileShare);
        }
        public void Dispose()
        {
            XMLFileStream.Close();
            Access.ReleaseMutex();
        }
    }
}
