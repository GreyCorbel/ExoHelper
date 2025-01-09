using System;
using System.Net;
namespace ExoHelper
{
    //exception class for ExoHelper module
    public class ExoException : Exception
    {
        public HttpStatusCode? StatusCode { get; set; }
        public string ExoErrorCode { get; set; }
        public string ExoErrorType { get; set; }
        public ExoException(HttpStatusCode? statusCode, string exoCode, string exoErrorType, string message):this(statusCode, exoCode, exoErrorType, message, null)
        {

        }
        public ExoException(HttpStatusCode? statusCode, string exoCode, string exoErrorType, string message, Exception innerException):base(message, innerException)
        {
            StatusCode = statusCode;
            ExoErrorCode = exoCode;
            ExoErrorType = exoErrorType;
        }
    }
}
