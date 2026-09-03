using System.Globalization;
using Avalonia.Data.Converters;
using Avalonia.Media;
using SharePointLinkManifestBuilder.Core.Models;

namespace SharePointLinkManifestBuilder.App.Converters;

/// <summary>
/// Compares a bound value with a parameter, for radio-button groups over enums.
/// Converting back returns the parameter when checked, so the group round-trips.
/// </summary>
public sealed class EnumEqualsConverter : IValueConverter
{
    /// <summary>The shared instance used from XAML.</summary>
    public static readonly EnumEqualsConverter Instance = new();

    /// <inheritdoc />
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value?.ToString()?.Equals(parameter?.ToString(), StringComparison.OrdinalIgnoreCase) == true;

    /// <inheritdoc />
    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture)
    {
        if (value is not true || parameter is null)
        {
            return Avalonia.Data.BindingOperations.DoNothing;
        }

        var enumType = Nullable.GetUnderlyingType(targetType) ?? targetType;

        return enumType.IsEnum && Enum.TryParse(enumType, parameter.ToString(), true, out var parsed)
            ? parsed
            : Avalonia.Data.BindingOperations.DoNothing;
    }
}

/// <summary>Inverts a boolean, for "enabled when not busy" bindings.</summary>
public sealed class InverseBooleanConverter : IValueConverter
{
    /// <summary>The shared instance used from XAML.</summary>
    public static readonly InverseBooleanConverter Instance = new();

    /// <inheritdoc />
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is not true;

    /// <inheritdoc />
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value is not true;
}

/// <summary>
/// Maps a link result status to a colour.
/// <para>
/// Colour is only ever a secondary cue: every status is also shown as text, because colour
/// alone excludes anyone with a colour-vision deficiency and disappears in high-contrast modes.
/// </para>
/// </summary>
public sealed class LinkStatusBrushConverter : IValueConverter
{
    /// <summary>The shared instance used from XAML.</summary>
    public static readonly LinkStatusBrushConverter Instance = new();

    /// <inheritdoc />
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value switch
        {
            LinkResultStatus.Created => Brushes.SeaGreen,
            LinkResultStatus.Reused or LinkResultStatus.Existing => Brushes.SteelBlue,
            LinkResultStatus.Skipped => Brushes.Gray,
            LinkResultStatus.PolicyBlocked => Brushes.DarkOrange,
            LinkResultStatus.AccessDenied => Brushes.IndianRed,
            LinkResultStatus.Unsupported => Brushes.DarkGoldenrod,
            LinkResultStatus.Failed => Brushes.IndianRed,
            _ => Brushes.Gray,
        };

    /// <inheritdoc />
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        Avalonia.Data.BindingOperations.DoNothing;
}

/// <summary>Formats a byte count for display.</summary>
public sealed class ByteSizeConverter : IValueConverter
{
    /// <summary>The shared instance used from XAML.</summary>
    public static readonly ByteSizeConverter Instance = new();

    /// <inheritdoc />
    public object Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value switch
        {
            long bytes and >= 1024L * 1024 * 1024 => $"{bytes / (1024.0 * 1024 * 1024):0.##} GB",
            long bytes and >= 1024 * 1024 => $"{bytes / (1024.0 * 1024):0.##} MB",
            long bytes and >= 1024 => $"{bytes / 1024.0:0.##} KB",
            long bytes => $"{bytes} B",
            _ => string.Empty,
        };

    /// <inheritdoc />
    public object ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        Avalonia.Data.BindingOperations.DoNothing;
}

/// <summary>
/// Bridges <see cref="DateTimeOffset"/> view models to <c>CalendarDatePicker.SelectedDate</c>,
/// which is <see cref="DateTime"/>.
/// <para>
/// Binding the two directly throws InvalidCastException at runtime -- the binding engine will
/// not narrow a DateTimeOffset to a DateTime for you. The view models keep DateTimeOffset
/// because that is what Microsoft Graph expects for a link expiry, and losing the offset on the
/// way to Graph would silently shift an expiry by the machine's UTC offset.
/// </para>
/// </summary>
public sealed class DateTimeOffsetToDateTimeConverter : IValueConverter
{
    /// <summary>The shared instance.</summary>
    public static readonly DateTimeOffsetToDateTimeConverter Instance = new();

    /// <inheritdoc />
    public object? Convert(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value switch
        {
            DateTimeOffset offset => offset.LocalDateTime,
            DateTime local => local,
            _ => null,
        };

    /// <inheritdoc />
    public object? ConvertBack(object? value, Type targetType, object? parameter, CultureInfo culture) =>
        value switch
        {
            // The picker yields a local wall-clock date with no offset. Attaching the machine's
            // current offset is what the user meant by picking that day where they are.
            DateTime local => new DateTimeOffset(DateTime.SpecifyKind(local, DateTimeKind.Local)),
            DateTimeOffset offset => offset,
            _ => null,
        };
}
