using System;

namespace AcGraphicToFrame.Exceptions
{
    internal class FormatNotFoundException : Exception
    {
        public FormatNotFoundException(string message)
            :base(message)
        {
        }
    }
}
