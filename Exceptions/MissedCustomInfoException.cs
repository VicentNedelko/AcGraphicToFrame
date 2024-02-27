using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
