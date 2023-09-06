using System;

namespace xjct
{
    /// <summary>
    /// Custom exception handling.
    /// </summary>
    class WhiteListException : Exception 
    {
        public WhiteListException() { }
        public WhiteListException(string message) { }
        public WhiteListException(string message, Exception iException) : base(message, iException) { }
    }


}
