using System.Messaging;

namespace DeltaCore.Queue
{
	public class MSMQ
	{
		public string QueuePath;
		private MessageQueue _queue;
		
		public MSMQ () {}

		public MSMQ(string queuePathConfig)
		{
			QueuePath = queuePathConfig;
		}

		public void Mensaje_Sender(object msgRequest)
		{
			var msg = new System.Messaging.Message {Body = msgRequest};
			_queue = MessageQueue.Exists(QueuePath) ? new MessageQueue(QueuePath) : MessageQueue.Create(QueuePath);
			_queue.Send(msg);
		}

		public void Mensaje_Sender(object msgRequest, string queuePath)
		{
			var msg = new System.Messaging.Message { Body = msgRequest };
			_queue = MessageQueue.Exists(queuePath) ? new MessageQueue(queuePath) : MessageQueue.Create(queuePath);
			_queue.Send(msg);
		}

		public object Mensaje_Receiver()
		{
			_queue = MessageQueue.Exists(QueuePath) ? new MessageQueue(QueuePath) : MessageQueue.Create(QueuePath);
			_queue.Formatter = new XmlMessageFormatter();
			var receive = _queue.Receive();
			return receive != null ? receive.Body : null;
		}

		public object Mensaje_Receiver(string queuePath)
		{
			_queue = MessageQueue.Exists(queuePath) ? new MessageQueue(queuePath) : MessageQueue.Create(queuePath);
			_queue.Formatter = new XmlMessageFormatter();
			var receive = _queue.Receive();
			return receive != null ? receive.Body : null;
		}
	}
}