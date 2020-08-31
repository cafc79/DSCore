using System;
using DeltaCore.Utilities.Research;
using MongoDB.Bson;
using MongoDB.Bson.IO;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using MongoDB.Driver.Linq;
using Newtonsoft.Json;

namespace DeltaCore.NoSQL
{
    public class MDBConexion
    {
        public IMongoClient _cnx;
        public IMongoDatabase _cmd;
        //public IMongoCollection<T> Collection { get; private set; }

        [Title("Default Settings")]
        [Category("1. Getting Server")]
        [Description("We can use MongoDB.Driver.MongoClient to create Client to interact with MongoDB. If we do not specify any arguement then it connects with mongodb instance running on localhost on default port [27017]")]
        [Code("Console", "mongo")]
        [Code("C#", "new MongoDB.Driver.MongoClient()\n\t.GetServer()\n\t.Connect()")]
        public void Connect(string connectionString, string database) {
            _cnx = new MongoClient(connectionString);
            _cmd = _cnx.GetDatabase(database);
            
        }

        public void CreateCollection(string Coleccion)
        {
            var options = new CreateCollectionOptions { Capped = true, MaxSize = 1024 * 1024 };
            _cmd.CreateCollection(Coleccion, options);
        }
        
        public void Insert(object clazz)
        {
            Object obj = Activator.CreateInstance(clazz.GetType());
            var collection = _cmd.GetCollection<BsonDocument>(obj.GetType().Name);
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(clazz, Formatting.Indented);
            BsonDocument BSONDoc = MongoDB.Bson.Serialization.BsonSerializer.Deserialize<BsonDocument>(json);
            collection.InsertOne(BSONDoc);
        }

        public void DeleteCustomer(object clazz, ObjectId objectId)
        {
            Object obj = Activator.CreateInstance(clazz.GetType());
            _cmd.DropCollection(obj.GetType().Name);
        }
        
        public BsonDocument GetInformacion(object clazz)
        {
            var lst = _cmd.ListCollectionNames();
            Object obj = Activator.CreateInstance(clazz.GetType());
            
            var collection = _cmd.GetCollection<BsonDocument>(obj.GetType().Name);

            var documents = collection.Find(new BsonDocument()).ToList();
            //var collection = _cnx.GetCollection<T>(typeof(T).Name.ToLower());
            //BsonDocument bsonCustomer = collection.FindOne(query);
            //Customer customer = BsonSerializer.Deserialize<Customer>(bsonCustomer);
            return collection.ToBsonDocument();
        }
    }
}