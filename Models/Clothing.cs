namespace Finch_Inventory.Models
{
    using System;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Data.Entity.Spatial;

    [Table("Clothing")]
    public partial class Clothing
    {
        public int ID { get; set; }

        public int PM_Number { get; set; }

        public int PositionID { get; set; }

        [StringLength(50)]
        public string Manufacturer { get; set; }

        public int TypeID { get; set; }

        [Required]
        [StringLength(50)]
        public string Serial_Number { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Received { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Placed_On_Mac { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Date_Removed_From_Mac { get; set; }

        public int StatusID { get; set; }

        public int? LocationID { get; set; }

        public int? Age { get; set; }

        [Column(TypeName = "date")]
        public DateTime? Today { get; set; }

        public virtual Location Location { get; set; }

        public virtual Position Position { get; set; }

        public virtual Status Status { get; set; }

        public virtual Type Type { get; set; }
    }
}
