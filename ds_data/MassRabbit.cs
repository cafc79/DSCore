using System;
using System.Threading.Tasks;
using MassTransit;
using MassTransit.RabbitMqTransport;
using RabbitMQ.Client;

namespace DeltaCore.Data
{
    public class MassRabbit
    {
        private Uri MsgServer;
        private string mqUsuario;
        private string mqPass;

        public MassRabbit(string RMQ, string Usuario, string Pass)
        {
            MsgServer = new Uri(RMQ);
            mqUsuario = Usuario;
            mqPass = Pass;            
        }

        public void RunMassTransitReceiverWithRabbit(string TopicFCM)
        {           
            ConnectionFactory factory = new ConnectionFactory();
            // "guest"/"guest" by default, limited to localhost connections
            factory.UserName = mqUsuario;
            factory.Password = mqPass;
            factory.HostName = MsgServer.AbsoluteUri;

            var busControl = Bus.Factory.CreateUsingRabbitMq(cfg =>
            {
                cfg.Host(MsgServer.AbsoluteUri, "/", h =>
                {
                    h.Username(mqUsuario);
                    h.Password(mqPass);
                });
                //registrationAction?.Invoke(cfg, h);
                cfg.ReceiveEndpoint(rabbitMqHost, TopicFCM, conf => { conf.Consumer<ConsumerFCM>(); });
            });
            busControl.Start();
            busControl.Stop();
        }
    }

    public class ConsumerFCM : IConsumer<IRegisterFCM>
    {
        public Task Consume(ConsumeContext<IRegisterFCM> context)
        {
            IRegisterFCM Notificacion = context.Message;
            return Task.FromResult(context.Message);
        }
    }

    public interface IRegisterFCM
    {
        Guid Id { get; }
        string Topic { get; }
        string Title { get; }
        string Body { get; }
    }
}
