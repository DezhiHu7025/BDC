using System;
 
 

namespace Kcis.Models
{

    public class KcisException : Exception
    {
        public KcisException(string strMessage)
            : base(strMessage)
        {
        }


    }


}