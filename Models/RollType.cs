namespace Finch_Inventory.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("RollType")]
    public partial class RollType
    {
        public int ID { get; set; }

        [Required]
        [StringLength(50)]
        public string Type { get; set; }
    }
}
