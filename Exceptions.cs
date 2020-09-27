using System;
using System.Text;

namespace OutlookWelkinSyncFunction
{
    public static class Exceptions
    {
        public static string ToStringRecursively(Exception exception)
        {
            StringBuilder stringBuilder = new StringBuilder();

            while (exception != null)
            {
                stringBuilder.AppendLine(exception.Message);
                stringBuilder.AppendLine(exception.StackTrace);

                exception = exception.InnerException;
            }

            return stringBuilder.ToString();
        }
    }
}