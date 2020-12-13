namespace OutlookWelkinSync
{
    using Newtonsoft.Json.Converters;

    public class JsonDateFormatConverter : IsoDateTimeConverter
{
    public JsonDateFormatConverter(string format)
    {
        DateTimeFormat = format;
    }
}
}