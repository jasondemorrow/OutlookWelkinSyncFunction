namespace OutlookWelkinSync
{
    using System;
    using Microsoft.Extensions.Caching.Memory;
    
    public class InMemoryCache<T>
    {
        private MemoryCache internalCache = new MemoryCache(new MemoryCacheOptions()
        {
            SizeLimit = 1024
        });
    
        public T GetOrCreate(object key, Func<T> factory)
        {
            T cacheEntry;
            if (!internalCache.TryGetValue(key, out cacheEntry))
            {
                cacheEntry = factory();
                MemoryCacheEntryOptions cacheEntryOptions = new MemoryCacheEntryOptions().SetSize(1);
                internalCache.Set(key, cacheEntry, cacheEntryOptions);
            }

            return cacheEntry;
        }
    }
}