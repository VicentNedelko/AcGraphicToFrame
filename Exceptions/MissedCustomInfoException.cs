using System;

namespace AcGraphicToFrame.Exceptions
{
    internal class MissedCustomInfoException : Exception
    {
        public MissedCustomInfoException(string message)
            :base(message)
        {
            
        }
    }
}
