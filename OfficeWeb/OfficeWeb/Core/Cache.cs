using Newtonsoft.Json.Linq;
using System;
using System.Runtime.Caching;

namespace OfficeWeb.Core
{
    /// <summary>
    /// 缓存框架
    /// </summary>
    public class CacheData
    {
        private const string CacheName = "System";
        private static MemoryCache MemoryCache = new MemoryCache(CacheName);
        /// <summary>
        /// 缓存过期时间
        /// </summary>
        private static DateTimeOffset DestroyTime { get { return DateTimeOffset.Now.AddMinutes(5); } }

        /// <summary>
        /// 设置缓存
        /// </summary>
        /// <param name="name">要插入的缓存项的唯一标识符</param>
        /// <param name="value">该缓存项的数据</param>
        public static void Set(string name, JObject value)
        {
            if (MemoryCache.Contains(name))
            {
                MemoryCache[name] = value;
            }
            else
            {
                MemoryCache.Set(name, value, DestroyTime);
            }
        }

        /// <summary>
        /// 从缓存中返回一个项
        /// </summary>
        /// <param name="name">要获取的缓存项的唯一标识符</param>
        /// <returns></returns>
        public static JObject Get(string name)
        {
            return MemoryCache.Contains(name) ? JObject.Parse(MemoryCache.Get(name).ToString()) : null;
        }

        /// <summary>
        /// 从缓存中移除某个缓存项
        /// </summary>
        /// <param name="name">要移除的缓存项的唯一标识符</param>
        public static void Remove(string name)
        {
            MemoryCache.Remove(name);
        }

        /// <summary>
        /// 确定缓存中是否包括指定的键值
        /// </summary>
        /// <param name="name">要搜索的缓存项的唯一标识符</param>
        /// <returns></returns>
        public static bool ContainsKey(string name)
        {
            return MemoryCache.Contains(name);
        }
    }
}
