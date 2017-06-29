namespace startUp
{
    using System;
    using System.Data.Entity;
    using System.Linq;
    using Models;

    public class OrganizationsContext : DbContext
    {
        public OrganizationsContext()
            : base("name=OrganizationsContext")
        {
        }

        public virtual DbSet<Organization> Organizations { get; set; }
    }

}