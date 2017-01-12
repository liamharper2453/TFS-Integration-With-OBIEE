namespace VSODataStructures_NS
{
    public class VSODataStructures_VSOTFSIndex
    {
        public int ID { get; set; }
        public int watermark { get; set; }

        public string title { get; set; }

    }

    public class VSODataStructures_VSODBIndex
    {
        public int ID { get; set; }
        public int watermark { get; set; }

        public string title { get; set; }

    }

    public class VSODataStructures_VSOItemSentToDB
    { 
    public int ID { get; set; }
    public string changedBy { get; set; }

    public string changedDate { get; set; }

    public string title { get; set; }
}


}
