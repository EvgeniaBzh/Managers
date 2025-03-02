using Microsoft.EntityFrameworkCore;

namespace WebApplication1.Models
{
    public class WebContext : DbContext
    {
        public WebContext(DbContextOptions<WebContext> options)
            : base(options)
        {
        }


        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            if (!optionsBuilder.IsConfigured)
            {
                optionsBuilder.UseSqlServer("Server=DESKTOP-DHK8L7H\\SQLEXPRESS;Database=Managers;Trusted_Connection=True;TrustServerCertificate=True;");
            }
        }
    }
}
