namespace DocMailer.Models
{
    /// <summary>
    /// Enum for different recipient filtering modes
    /// </summary>
    public enum SendMode
    {
        All,
        NotSent,
        NotResponded,
        Test,
        Specific
    }
}
