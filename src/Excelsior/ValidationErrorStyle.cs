namespace Excelsior;

/// <summary>
/// Controls how Excel responds when a cell value fails data validation.
/// </summary>
public enum ValidationErrorStyle
{
    /// <summary>
    /// Block the invalid entry. The user must enter a valid value or cancel.
    /// </summary>
    Stop,

    /// <summary>
    /// Warn the user and let them choose to keep the invalid value.
    /// </summary>
    Warning,

    /// <summary>
    /// Inform the user; the value is accepted regardless.
    /// </summary>
    Information
}
