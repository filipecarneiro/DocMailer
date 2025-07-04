using System.ComponentModel.DataAnnotations;

namespace DocMailer.Models
{
    /// <summary>
    /// Represents a recipient/subscriber
    /// </summary>
    public class Recipient
    {
        [Required]
        public string Name { get; set; } = string.Empty;
        
        [Required]
        [EmailAddress]
        public string Email { get; set; } = string.Empty;
        
        public string Company { get; set; } = string.Empty;
        
        public string Position { get; set; } = string.Empty;
        
        public string FirstName { get; set; } = string.Empty;
        
        public DateTime SubscriptionDate { get; set; } = DateTime.Now;
        
        public DateTime? LastSent { get; set; }
        
        public bool? Responded { get; set; }
        
        public bool IsCanceled { get; set; } = false;
        
        public string LastSentStatus { get; set; } = string.Empty;
        
        public int RowNumber { get; set; } // To track Excel row for updates
        
        /// <summary>
        /// Gets the first name - uses explicit FirstName if available, otherwise extracts from Name field
        /// </summary>
        public string GetFirstName() => !string.IsNullOrEmpty(FirstName) ? FirstName.Trim() : Name?.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim() ?? "";
        
        public Dictionary<string, object> CustomFields { get; set; } = new Dictionary<string, object>();
    }
}
