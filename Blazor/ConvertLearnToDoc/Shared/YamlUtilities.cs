using YamlDotNet.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.EventEmitters;

namespace ConvertLearnToDoc.Shared;

public static class YamlUtilities
{
    public static string ObjectToYamlString(object input)
    {
        return new SerializerBuilder()
            //.WithNamingConvention(CamelCaseNamingConvention.Instance)
            .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
            .WithEventEmitter(nextEmitter => new MultilineQuoteFixupEmitter(nextEmitter))
            .Build()
            .Serialize(input);
    }

    public static string YamlStringToJson(string text)
    {
        var serializer = new SerializerBuilder().JsonCompatible().Build();
        return serializer.Serialize(text);
    }

    public static string DictionaryToYamlString(Dictionary<object, object> values)
    {
        return new SerializerBuilder()
            //.WithNamingConvention(CamelCaseNamingConvention.Instance)
            .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
            .WithEventEmitter(nextEmitter => new MultilineQuoteFixupEmitter(nextEmitter))
            .Build()
            .Serialize(values);
    }

    public static Dictionary<object, object>? YamlStringToDictionary(string yamlText)
    {
        return new Deserializer()
                .Deserialize(new StringReader(yamlText))
            as Dictionary<object, object>;
    }

    sealed class MultilineQuoteFixupEmitter : ChainedEventEmitter
    {
        public MultilineQuoteFixupEmitter(IEventEmitter nextEmitter)
            : base(nextEmitter) { }

        public override void Emit(ScalarEventInfo eventInfo, IEmitter emitter)
        {
            if (typeof(string).IsAssignableFrom(eventInfo.Source.Type))
            {
                var value = eventInfo.Source.Value as string;
                if (!string.IsNullOrEmpty(value))
                {
                    var isMultiLine = value.IndexOfAny(new char[] { '\r', '\n', '\x85', '\x2028', '\x2029' }) >= 0;
                    if (isMultiLine)
                    {
                        eventInfo = new ScalarEventInfo(eventInfo.Source)
                        {
                            Style = ScalarStyle.Literal
                        };
                    }
                    else
                    {
                        if (value.Contains(':'))
                            eventInfo.Style = ScalarStyle.DoubleQuoted;
                    }
                }
            }

            base.Emit(eventInfo, emitter);
        }
    }
}