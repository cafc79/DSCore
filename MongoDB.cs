using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MongoDB.Bson;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Driver;
using MongoDB.Driver.Builders;
using MongoDB.Driver.Linq;


namespace DeltaCore.NoSQL
{
		public class MongoDB
		{
				public class MongoDB
				{
						public void Connect()
						{
								// Create server settings to pass connection string, timeout, etc.
								MongoServerSettings settings = new MongoServerSettings();
								settings.Server = new MongoServerAddress("localhost", 27017);
								// Create server object to communicate with our server
								MongoServer server = new MongoServer(settings);
								// Get our database instance to reach collections and data
								var database = server.GetDatabase("MessageDB");
						}
				}

				private MongoDatabase Database
{  get
   {
     string connectionString = "mongodb://localhost";
     MongoClient client = new MongoClient(connectionString);
     MongoServer server = client.GetServer();
     MongoDatabase database = server.GetDatabase("test");
     return database;

  }

						public void CreateCustomer(Customer customer)
{  var collection = Database.GetCollection("Customers");
  collection.Insert(customer);
}

　
public Customer GetCustomer(ObjectId objectid
{
 //Query: Get Customer.id == supported objectid
 IMongoQuery query = Query<Customer>.EQ(e => e.Id, objectid);  
 var collection = Database.GetCollection("Customers");
 BsonDocument bsonCustomer = collection.FindOne(query);
 Customer customer = BsonSerializer.Deserialize<Customer>(bsonCustomer);

 return customer;
 }

　
  public void DeleteCustomer(ObjectId objectId)
{
 var query = Query<Customer>.EQ(e => e.Id, objectId);
 var collection = Database.GetCollection("Customers");

 collection.Remove(query);

}

　
public List<Customer> GetAllCustomer()
{  var collection = Database.GetCollection("Customers");
  MongoCursor<BsonDocument> Bsoncustomers = collection.FindAll();
  List<Customer> customers = new List<Customer>();
  foreach (var bsoncustomer in Bsoncustomers)
  {    var cus = BsonSerializer.Deserialize<Customer>(bsoncustomer);
    customers.Add(cus);  
  }
}

 public List<Customer> GetByAddress6(string address)
{ IMongoQuery query = Query<Customer>.EQ(x => x.Addresses[0].Address6 == address);
 return GetByQuery(query);
}

　
public List<Customer> GetByQuery(IMongoQuery query)
{ var collection = Database.GetCollection("Customers");
 MongoCursor<BsonDocument> Bsoncustomers = collection.Find(query);
 List<Customer> customers = new List<Customer>();
 foreach (var bsoncustomer in Bsoncustomers)
 {   var cus = BsonSerializer.Deserialize<Customer>(bsoncustomer);
   customers.Add(cus);  
 }
  return customers;
 }
}
}  

				public void getReader
						//Builds new Collection if 'entities' is not found
MongoCollection<ClubMember> collection = database.GetCollection<ClubMember>("entities");

Console.WriteLine("List of ClubMembers in collection ...");
MongoCursor<ClubMember> members = collection.FindAll();
foreach (ClubMember clubMember in members)
{
    clubMember.PrintDetailsToScreen();
}


		//Build an index if it is not already built
IndexKeysBuilder keys = IndexKeys.Ascending("Lastname", "Forename").Descending("Age");

//Add an optional name- useful for admin
IndexOptionsBuilder options = IndexOptions.SetName("myIndex");

//This locks the database while the index is being built
collection.EnsureIndex(keys, options);


		var names =
    collection.AsQueryable().Where(p => p.Lastname.StartsWith("R") && p.Forename.EndsWith("an")).OrderBy(
        p => p.Lastname).ThenBy(p => p.Forename).Select(p => new { p.Forename, p.Lastname });

Console.WriteLine("Members where the Lastname starts with 'R' and the Forename ends with 'an'");
foreach (var name in names)
{
    Console.WriteLine(name.Lastname + " " + name.Forename);
}

var regex = new Regex("ar");
Console.WriteLine("List of Lastnames containing the substring 'ar'");
IQueryable<string> regexquery =
    collection.AsQueryable().Where(py => regex.IsMatch(py.Lastname)).Select(p => p.Lastname).Distinct();

foreach (string name in regexquery)
{
    Console.WriteLine(name);
}
		}
}
