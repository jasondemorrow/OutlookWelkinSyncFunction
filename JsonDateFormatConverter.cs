using Newtonsoft.Json.Converters;

namespace OutlookWelkinSyncFunction
{
    public class JsonDateFormatConverter : IsoDateTimeConverter
{
    public JsonDateFormatConverter(string format)
    {
        DateTimeFormat = format;
    }
}
}