using System;
using System.Timers;

namespace GoogleDriveFileUploaded
{
    internal class timer
    {
        internal bool autoreset;

        public timer(int v)
        {
        }

        public Action<object, ElapsedEventArgs> elapsed { get; internal set; }
        public bool enabled { get; internal set; }
    }
}