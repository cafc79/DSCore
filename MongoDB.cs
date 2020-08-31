using System;

using MongoDB.Bson;
using MongoDB.Bson.IO;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using MongoDB.Driver.Linq;
using Newtonsoft.Json;

namespace DeltaCore.NoSQL
{
    public class MongoDB<T> where T : class
    {
        protected static IMongoClient _cmd;
        protected static IMongoDatabase _cnx;
        public IMongoCollection<T> Collection { get; private set; }

        private void Connect(string connectionString, string database) {
            _cmd = new MongoClient(connectionString);
            _cnx = _cmd.GetDatabase(database);
        }

        public void CreateCollection(string Coleccion)
        {
            var options = new CreateCollectionOptions { Capped = true, MaxSize = 1024 * 1024 };
            _cnx.CreateCollection(Coleccion, options);
        }
        
        public void Insert(object clazz)
        {
            Object obj = Activator.CreateInstance(clazz.GetType());
            var collection = _cnx.GetCollection<BsonDocument>(obj.GetType().Name);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(clazz, Formatting.Indented);
            BsonDocument BSONDoc = MongoDB.Bson.Serialization.BsonSerializer.Deserialize<BsonDocument>(json);
            collection.InsertOne(BSONDoc);
        }

        public void DeleteCustomer(object clazz, ObjectId objectId)
        {
            Object obj = Activator.CreateInstance(clazz.GetType());
            _cnx.DropCollection(obj.GetType().Name);
        }
        
        public BsonDocument GetInformacion()
        {
            var collection = _cnx.GetCollection<T>(typeof(T).Name.ToLower());
            //BsonDocument bsonCustomer = collection.FindOne(query);
            //Customer customer = BsonSerializer.Deserialize<Customer>(bsonCustomer);
            return collection.ToBsonDocument();
        }
    }
}