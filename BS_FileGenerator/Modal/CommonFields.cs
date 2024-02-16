namespace BS_FileGenerator.Modal
{
    public class CommonFields
    {
        public int CompanyId { get; set; }

        public string CompanyName { get; set; }

        public string DeveloperName { get; set; }

        public List<string> ToAddress { get; set; }

        public string Body { get; set; }

        public KafkaServiceName kafkaServiceName { get; set; }

        public string LocalConnectionString { get; set; }
    }
}
