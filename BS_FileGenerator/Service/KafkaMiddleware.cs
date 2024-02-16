using Bot.CoreBottomHalf.CommonModal.HtmlTemplateModel;
using Bot.CoreBottomHalf.CommonModal.Kafka;
using Confluent.Kafka;
using CoreBottomHalf.CommonModal.HtmlTemplateModel;
using Microsoft.Extensions.Options;
using ModalLayer;
using Newtonsoft.Json;

namespace EmailRequest.Service
{
    public class KafkaService : IHostedService
    {
        private readonly KafkaServiceConfig _kafkaServiceConfig;
        private ILogger<KafkaService> _logger;
        private IServiceProvider _serviceProvider;

        public KafkaService(IServiceProvider serviceProvider, ILogger<KafkaService> logger, IOptions<KafkaServiceConfig> options)
        {
            _serviceProvider = serviceProvider;
            _kafkaServiceConfig = options.Value;
            _logger = logger;
        }

        public KafkaServiceConfig KafkaServiceConfig => _kafkaServiceConfig;

        public Task StartAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("[Kafka] Kafka listener registered successfully.");

            Task.Run(() =>
            {
                SubscribeKafkaTopic();
            });

            return Task.CompletedTask;
        }

        public Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Stoping service.");
            return Task.CompletedTask;
        }

        private void SubscribeKafkaTopic()
        {
            using (IServiceScope scope = _serviceProvider.CreateScope())
            {
                var config = new ConsumerConfig
                {
                    GroupId = "gid-consumers",
                    BootstrapServers = $"{KafkaServiceConfig.ServiceName}:{KafkaServiceConfig.Port}"
                };

                _logger.LogInformation($"[Kafka] Start listning kafka topic: {KafkaServiceConfig.Topic}");
                using (var consumer = new ConsumerBuilder<Null, string>(config).Build())
                {
                    consumer.Subscribe(KafkaServiceConfig.Topic);
                    while (true)
                    {
                        try
                        {
                            _logger.LogInformation($"[Kafka] Waiting on topic: {KafkaServiceConfig.Topic}");
                            var message = consumer.Consume();

                            HandleMessageSendEmail(message, scope);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError($"[Kafka Error]: Got exception - {ex.Message}");
                        }
                    }
                }
            }
        }

        private void HandleMessageSendEmail(ConsumeResult<Null, string> result, IServiceScope scope)
        {
            if (string.IsNullOrWhiteSpace(result.Message.Value))
                throw new Exception("[Kafka] Received invalid object from producer.");

            _logger.LogInformation($"[Kafka] Message received: {result.Message.Value}");


            KafkaPayload commonFields = JsonConvert.DeserializeObject<KafkaPayload>(result.Message.Value);
            if (commonFields == null)
                throw new Exception("[Kafka] Received invalid object from producer.");

            switch (commonFields.kafkaServiceName)
            {
                case KafkaServiceName.Billing:
                    BillingTemplateModel billingTemplateModel = JsonConvert.DeserializeObject<BillingTemplateModel>(result.Message.Value);
                    if (billingTemplateModel == null)
                        throw new Exception("[Kafka] Received invalid object for billing template modal from producer.");
                    break;
            }
        }
    }
}
