    using System;
    using Microsoft.EntityFrameworkCore;
    using Microsoft.EntityFrameworkCore.Metadata;

    namespace Repository.Models
    {
        public partial class Context : DbContext
        {
            public Context()
            {
            }

            public Context(DbContextOptions<Context> options)
                : base(options)
            {
            }
            //Constructor used int test cases for passing conncetion string
            //==============Start
            public Context(string connectionString) : base(GetOptions(connectionString))
            {
            }
            private static DbContextOptions GetOptions(string connectionString)
            {
                return SqlServerDbContextOptionsExtensions.UseSqlServer(new DbContextOptionsBuilder(), connectionString).Options;
            }
            //==============End

            public virtual DbSet<Login> tblUserDetails { get; set; }
            public virtual DbSet<Operations> Operations { get; set; }
            public virtual DbSet<WorkCenterMaster> WorkCenterMaster { get; set; }


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
            {
                if (!optionsBuilder.IsConfigured)
                {
                    #warning To protect potentially sensitive information in your connection string, you should move it out of source code. See http://go.microsoft.com/fwlink/?LinkId=723263 for guidance on storing connection strings.
                    optionsBuilder.UseSqlServer("Server=127.0.0.1;Database=master;User ID=sa;Password=Strong.Pwd-123;");
                }
            }

            protected override void OnModelCreating(ModelBuilder modelBuilder)
            {
                modelBuilder.HasAnnotation("ProductVersion", "2.2.6-servicing-10079");
             
            }
        }
    }
