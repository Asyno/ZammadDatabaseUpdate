using System.Runtime.Serialization;

namespace ZammadDatabaseUpdate
{
    /// <summary>
    /// Data Contract for the Zammad User
    /// </summary>
    [DataContract]
    internal class User
    {
        [DataMember(Name = "id")]
        internal string id { get; set; }
        
        [DataMember(Name = "firstname")]
        internal string firstname { get; set; }

        [DataMember(Name = "lastname")]
        internal string lastname { get; set; }

        [DataMember(Name = "support-level")]
        internal string support_level { get; set; }

        [DataMember(Name = "support")]
        internal string support { get; set; }

        [DataMember(Name = "product")]
        internal string product { get; set; }

        internal bool exist { get; set; }
    }
}
