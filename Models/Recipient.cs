using System.ComponentModel.DataAnnotations;

namespace DocMailer.Models
{
    /// <summary>
    /// Represents a recipient/subscriber
    /// </summary>
    public class Recipient
    {
        [Required]
        public string DisplayName { get; set; } = string.Empty;
        
        [Required]
        [EmailAddress]
        public string Email { get; set; } = string.Empty;
        
        public string Company { get; set; } = string.Empty;
        
        public string Position { get; set; } = string.Empty;
        
        public string FullName { get; set; } = string.Empty;
        
        public DateTime SubscriptionDate { get; set; } = DateTime.Now;
        
        public DateTime? LastSent { get; set; }
        
        public bool? Responded { get; set; }
        
        public bool IsCanceled { get; set; } = false;
        
        public string LastSentStatus { get; set; } = string.Empty;
        
        public int RowNumber { get; set; } // To track Excel row for updates
        
        /// <summary>
        /// Computed first name - uses DisplayName if available, then FullName, then extracts from DisplayName field
        /// </summary>
        public string FirstName => GetFirstName();
        
        /// <summary>
        /// Gets the first name - uses explicit FullName if available, otherwise extracts from DisplayName field
        /// </summary>
        public string GetFirstName() => !string.IsNullOrEmpty(FullName) ? FullName.Trim() : DisplayName?.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim() ?? "";
        
        public Dictionary<string, object> CustomFields { get; set; } = new Dictionary<string, object>();
    }
}
