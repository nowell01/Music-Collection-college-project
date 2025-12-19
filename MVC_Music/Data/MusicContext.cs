using Microsoft.EntityFrameworkCore;
using MVC_Music.Models;
using System.Numerics;
using MVC_Music.ViewModels;

namespace MVC_Music.Data
{
    public class MusicContext : DbContext
    {
        //To give access to IHttpContextAccessor for Audit Data with IAuditable
        private readonly IHttpContextAccessor _httpContextAccessor;

        //Property to hold the UserName value
        public string UserName
        {
            get; private set;
        }

        public MusicContext(DbContextOptions<MusicContext> options, IHttpContextAccessor httpContextAccessor)
            : base(options)
        {
            _httpContextAccessor = httpContextAccessor;
            if (_httpContextAccessor.HttpContext != null)
            {
                //We have a HttpContext, but there might not be anyone Authenticated
                UserName = _httpContextAccessor.HttpContext?.User.Identity.Name;
                UserName ??= "Unknown";
            }
            else
            {
                //No HttpContext so seeding data
                UserName = "Seed Data";
            }

        }

        public MusicContext(DbContextOptions<MusicContext> options)
            : base(options)
        {
            _httpContextAccessor = null!;
            UserName = "Seed Data";
        }


        public DbSet<Instrument> Instruments { get; set; }
        public DbSet<Musician> Musicians { get; set; }
        public DbSet<Play> Plays { get; set; }
        public DbSet<Genre> Genres { get; set; }
        public DbSet<Album> Albums { get; set; }
        public DbSet<Song> Songs { get; set; }
        public DbSet<Performance> Performances { get; set; }
        public DbSet<MusicianDocument> MusicianDocuments { get; set; }
        public DbSet<UploadedFile> UploadedFiles { get; set; }

        public DbSet<MusicianPhoto> MusicianPhotos { get; set; }
        public DbSet<MusicianThumbnail> MusicianThumbnails { get; set; }

        public DbSet<PerformanceReportVM> PerformanceReports { get; set; }


        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            //Prevent Cascade Delete from Instrument to Musician
            //so we are prevented from deleting a Instrument that is
            //the primary instrument of any Musician
            modelBuilder.Entity<Instrument>()
                .HasMany<Musician>(d => d.Musicians)
                .WithOne(p => p.Instrument)
                .HasForeignKey(p => p.InstrumentID)
                .OnDelete(DeleteBehavior.Restrict);

            //Prevent Cascade Delete from Instrument to Play (Parent Perspective)
            //so we are prevented from deleting a Instrument that
            //any Musician Plays
            modelBuilder.Entity<Instrument>()
                .HasMany<Play>(p => p.Plays)
                .WithOne(c => c.Instrument)
                .HasForeignKey(c => c.InstrumentID)
                .OnDelete(DeleteBehavior.Restrict);

            //Other Cascade Delete Restrictions
            //Note: since GenreID is a nullable foreign key in
            //  Song, we do not need to add any code to get the
            //  Cascade Delete Restriction between them.
            //Genre to Album
            modelBuilder.Entity<Genre>()
                .HasMany<Album>(p => p.Albums)
                .WithOne(c => c.Genre)
                .HasForeignKey(c => c.GenreID)
                .OnDelete(DeleteBehavior.Restrict);
            //Album to song
            modelBuilder.Entity<Album>()
                .HasMany<Song>(p => p.Songs)
                .WithOne(c => c.Album)
                .HasForeignKey(c => c.AlbumID)
                .OnDelete(DeleteBehavior.Restrict);
            //Instrument to Performance
            modelBuilder.Entity<Instrument>()
                .HasMany<Performance>(p => p.Performances)
                .WithOne(c => c.Instrument)
                .HasForeignKey(c => c.InstrumentID)
                .OnDelete(DeleteBehavior.Restrict);
            //Musician to Performance
            modelBuilder.Entity<Musician>()
                .HasMany<Performance>(p => p.Performances)
                .WithOne(c => c.Musician)
                .HasForeignKey(c => c.MusicianID)
                .OnDelete(DeleteBehavior.Restrict);
            //Note: We will allow the default Cascade Delete between Song and Performance

            //Many to Many Primary Key
            modelBuilder.Entity<Play>()
            .HasKey(p => new { p.MusicianID, p.InstrumentID });

            //Add a unique index to the SIN Number
            modelBuilder.Entity<Musician>()
            .HasIndex(p => p.SIN)
            .IsUnique();

            //Add a unique index to the Song Title on a Album
            modelBuilder.Entity<Song>()
            .HasIndex(p => new { p.Title, p.AlbumID })
            .IsUnique();

            //Add a unique composite index to the Performance
            modelBuilder.Entity<Performance>()
            .HasIndex(p => new { p.SongID, p.MusicianID, p.InstrumentID })
            .IsUnique();

            modelBuilder
                .Entity<PerformanceReportVM>()
                .ToView(nameof(PerformanceReports))
                .HasKey(v => v.ID);

            //Add a unique index to the Instrument Name
            modelBuilder.Entity<Instrument>()
            .HasIndex(p => p.Name)
            .IsUnique();
        }

        public override int SaveChanges(bool acceptAllChangesOnSuccess)
        {
            OnBeforeSaving();
            return base.SaveChanges(acceptAllChangesOnSuccess);
        }

        public override Task<int> SaveChangesAsync(bool acceptAllChangesOnSuccess, CancellationToken cancellationToken = default(CancellationToken))
        {
            OnBeforeSaving();
            return base.SaveChangesAsync(acceptAllChangesOnSuccess, cancellationToken);
        }

        private void OnBeforeSaving()
        {
            var entries = ChangeTracker.Entries();
            foreach (var entry in entries)
            {
                if (entry.Entity is IAuditable trackable)
                {
                    var now = DateTime.UtcNow;
                    switch (entry.State)
                    {
                        case EntityState.Modified:
                            trackable.UpdatedOn = now;
                            trackable.UpdatedBy = UserName;
                            break;

                        case EntityState.Added:
                            trackable.CreatedOn = now;
                            trackable.CreatedBy = UserName;
                            trackable.UpdatedOn = now;
                            trackable.UpdatedBy = UserName;
                            break;
                    }
                }
            }
        }


    }
}
