using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace RentSync.Core.Services;

/// <summary>
/// Reads/writes <see cref="DateOnly"/> as an ISO-8601 date string ("yyyy-MM-dd").
///
/// On .NET Framework, DateOnly comes from the Portable.System.DateTimeOnly
/// polyfill, which supplies the type but NOT a System.Text.Json converter, so
/// deserialising "2022-12-01" into DateOnly fails without this. Registering the
/// converter explicitly makes the same code work on net48 and net10.0 alike.
/// </summary>
public sealed class DateOnlyJsonConverter : JsonConverter<DateOnly>
{
    private const string Format = "yyyy-MM-dd";

    public override DateOnly Read(ref Utf8JsonReader reader, Type typeToConvert,
        JsonSerializerOptions options)
    {
        var text = reader.GetString();
        return DateOnly.ParseExact(text!, Format, CultureInfo.InvariantCulture);
    }

    public override void Write(Utf8JsonWriter writer, DateOnly value,
        JsonSerializerOptions options)
    {
        writer.WriteStringValue(value.ToString(Format, CultureInfo.InvariantCulture));
    }
}
