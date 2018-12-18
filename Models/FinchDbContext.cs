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

        public virtual DbSet<Clothing> Clothing { get; set; }
        public virtual DbSet<Location> Location { get; set; }
        public virtual DbSet<Machines> Machines { get; set; }
        public virtual DbSet<Position> Position { get; set; }
        public virtual DbSet<RollType> RollType { get; set; }
        public virtual DbSet<Status> Status { get; set; }
        public virtual DbSet<Type> Type { get; set; }
        public virtual DbSet<UserRoles> UserRoles { get; set; }
        public virtual DbSet<Users> Users { get; set; }

        protected override void OnModelCreating(DbModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Machines>()
                .HasMany(e => e.Clothing)
                .WithRequired(e => e.Machines)
                .HasForeignKey(e => e.PM_Number)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Position>()
                .HasMany(e => e.Clothing)
                .WithRequired(e => e.Position)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Status>()
                .HasMany(e => e.Clothing)
                .WithRequired(e => e.Status)
                .WillCascadeOnDelete(false);

            modelBuilder.Entity<Type>()
                .HasMany(e => e.Clothing)
                .WithRequired(e => e.Type)
                .WillCascadeOnDelete(false);
        }
    }
}
