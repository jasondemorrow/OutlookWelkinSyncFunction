using System;

namespace OutlookWelkinSyncFunction
{
    public static class Throw
    {
        public static void IfAnyAreNull(params object[] args)
        {
            if (args != null && args.Length > 0)
            {
                foreach (object obj in args)
                {
                    if (obj == null)
                    {
                        throw new ArgumentException($"Null parameter: {nameof(obj)}");
                    }
                }
            }
        }
    }
}