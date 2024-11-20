using System.ComponentModel.DataAnnotations;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json.Serialization;
using YamlDotNet.Core;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.EventEmitters;
using YamlDotNet.Serialization.NamingConventions;

namespace ConvertLearnToDoc.Shared;

public static class PersistenceUtilities
{
    public static string ObjectToJsonString(object obj)
    {
        return JsonConvert.SerializeObject(obj, Formatting.None,
            new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                NullValueHandling = NullValueHandling.Ignore
            });
    }

    public static T? JsonStringToObject<T>(string jsonText) where T : class, new()
    {
        return JsonConvert.DeserializeObject<T>(jsonText,
            new JsonSerializerSettings
            {
                ContractResolver = new CamelCasePropertyNamesContractResolver(),
                NullValueHandling = NullValueHandling.Ignore,
                DateFormatString = "MM/dd/yyyy" // 06/21/2021
            });
    }

    public static Dictionary<object, object>? JsonStringToDictionary(string jsonText)
    {
        var result = JsonStringToObject<Dictionary<object, object>>(jsonText);
        if (result == null)
            return null;

        var converted = new Dictionary<object, object>();
        foreach (var (key, value) in result)
        {
            if (value is JArray array)
            {
                // Convert JArray to List<string> - we assume simple text values only.
                converted.Add(key, array.Select(t => t.ToString()).ToList());
            }
            else
            {
                converted.Add(key, value ?? "");
            }

        }

        return converted;
    }

    public static string ObjectToYamlString(object input)
    {
        var result = new SerializerBuilder()
            .WithNamingConvention(CamelCaseNamingConvention.Instance)
            .ConfigureDefaultValuesHandling(DefaultValuesHandling.OmitNull)
            .WithEventEmitter(nextEmitter => new MultilineQuoteFixupEmitter(nextEmitter))
            .Build()
            .Serialize(input);

        return result;
    }

    //public static string YamlStringToJson(string text)
    //{
    //    var serializer = new SerializerBuilder().JsonCompatible().Build();
    //    return serializer.Serialize(text);
    //}

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
                    var isMultiLine = value.IndexOfAny(new[] { '\r', '\n', '\x85', '\x2028', '\x2029' }) >= 0;
                    if (isMultiLine)
                    {
                        if (value.TrimStart().StartsWith('-')) // sequence
                        {
                            eventInfo = new ScalarEventInfo(eventInfo.Source)
                            {
                                Style = ScalarStyle.Folded
                            };
                        }
                        else
                        {
                            eventInfo = new ScalarEventInfo(eventInfo.Source)
                            {
                                Style = ScalarStyle.Literal
                            };
                        }
                    }
                    else
                    {
                        char singleQuote = '\'';
                        char doubleQuote = '"';

                        bool isSingleQuoted = value.StartsWith(singleQuote) && value.EndsWith(singleQuote);
                        bool isDoubleQuoted = !isSingleQuoted && value.StartsWith(doubleQuote) && value.EndsWith(doubleQuote);
                        bool isQuoted = isSingleQuoted || isDoubleQuoted;

                        if (value.Contains(':') && !isQuoted)
                        {
                            eventInfo.Style = value.Contains(singleQuote) 
                                ? ScalarStyle.DoubleQuoted 
                                : ScalarStyle.SingleQuoted;
                        }

                        // If the existing value was already quoted (start/end)
                        // Then replace the value without quotes so we don't end up double-quoting.
                        // This is useful in the case where authors quote their values during editing.
                        if (isQuoted)
                        {
                            var cs = eventInfo.Source;
                            eventInfo = new ScalarEventInfo(
                                new ObjectDescriptor(value.Substring(1, value.Length - 2), cs.Type, cs.StaticType))
                                {
                                    Style = isSingleQuoted
                                        ? ScalarStyle.SingleQuoted
                                        : ScalarStyle.DoubleQuoted
                                };
                        }
                    }

                }
            }
            base.Emit(eventInfo, emitter);
        }
    }
}