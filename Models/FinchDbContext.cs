namespace Finch_Inventory.Models
{
    using System;
    using System.Data.Entity;
    using System.ComponentModel.DataAnnotations.Schema;
    using System.Linq;

    public partial class FinchDbContext : DbContext
    {
        public FinchDbContext()
            : base("name=FinchDbContext")
        {
        }

        public virtual DbSet<Clothing> Clothings { get; set; }
        public virtual DbSet<Location> Locations { get; set; }
        public virtual DbSet<Position> Positions { get; set; }
        public virtual DbSet<Status> Status { get; set; }
        public virtual DbSet<Type> Types { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Position>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Position)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Status>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Status)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Type>()
                .HasMany(e => e.Clothings)
                .WithRequired(e => e.Type)
                .WillCascadeOnDelete(false);
        }
    }
}
