using System;

namespace AcGraphicToFrame.Exceptions
{
    internal class AttributesNotFoundException : Exception
    {
        public AttributesNotFoundException(string message)
            :base(message)
        {
        }
    }
}
