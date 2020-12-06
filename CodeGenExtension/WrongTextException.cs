using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeGenExtension
{
    public class WrongTextException : Exception
    {
        public WrongTextException(string message) : base(message)
        {
        }
    }
}
