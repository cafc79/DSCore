using System;
using System.Collections.Generic;
using ServiceStack.Redis;
using System.Threading.Tasks;

namespace DeltaCore.Data
{
    public class Redis
    {
        RedisClient redis;

        public Redis(string redisURL)
        {
            redis = new RedisClient(redisURL, 6379);
        }
        
        public void Set(KeyValuePair<string, object> kv)
        {
            redis.SetEntry(kv.Key, kv.Value.ToString());
        }

        public string Get(string key)
        {
            return redis.GetValue(key);
        }
    }
}
