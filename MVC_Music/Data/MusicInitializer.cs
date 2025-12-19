using Humanizer.Localisation;
using Microsoft.CodeAnalysis;
using Microsoft.EntityFrameworkCore;
using MVC_Music.Models;
using MVC_Music.Utilities;
using System.Diagnostics;

namespace MVC_Music.Data
{
    public static class MusicInitializer
    {
        /// <summary>
        /// Prepares the Database and seeds data as required
        /// </summary>
        /// <param name="serviceProvider">DI Container</param>
        /// <param name="DeleteDatabase">Delete the database and start from scratch</param>
        /// <param name="UseMigrations">Use Migrations or EnsureCreated</param>
        /// <param name="SeedSampleData">Add optional sample data</param>
        public static void Initialize(IServiceProvider serviceProvider,
            bool DeleteDatabase = false, bool UseMigrations = true, bool SeedSampleData = true)
        {

            using (var context = new MusicContext(
                serviceProvider.GetRequiredService<DbContextOptions<MusicContext>>()))
            {
                //Refresh the database as per the parameter options
                #region Prepare the Database
                try
                {
                    //Note: .CanConnect() will return false if the database is not there!
                    if (DeleteDatabase || !context.Database.CanConnect())
                    {
                        if (!SqLiteDBUtility.ReallyEnsureDeleted(context)) //Delete the existing version 
                        {
                            Debug.WriteLine("Could not clear the old version " +
                                "of the database out of the way.  You will need to exit " +
                                "Visual Studio and try to do it manually.");
                        }
                        if (UseMigrations)
                        {
                            context.Database.Migrate(); //Create the Database and apply all migrations
                        }
                        else
                        {
                            context.Database.EnsureCreated(); //Create and update the database as per the Model
                        }
                        //Here is a good place to create any additional database objects such as Triggers or Views
                        //----------------------------------------------------------------------------------------
                        //Create the Triggers
                        string sqlCmd = @"
                            CREATE TRIGGER SetMusicianTimestampOnUpdate
                            AFTER UPDATE ON Musicians
                            BEGIN
                                UPDATE Musicians
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END;
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                            CREATE TRIGGER SetMusicianTimestampOnInsert
                            AFTER INSERT ON Musicians
                            BEGIN
                                UPDATE Musicians
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                            CREATE TRIGGER SetAlbumTimestampOnUpdate
                            AFTER UPDATE ON Albums
                            BEGIN
                                UPDATE Albums
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END;
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                            CREATE TRIGGER SetAlbumTimestampOnInsert
                            AFTER INSERT ON Albums
                            BEGIN
                                UPDATE Albums
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                            CREATE TRIGGER SetSongTimestampOnUpdate
                            AFTER UPDATE ON Songs
                            BEGIN
                                UPDATE Songs
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END;
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                            CREATE TRIGGER SetSongTimestampOnInsert
                            AFTER INSERT ON Songs
                            BEGIN
                                UPDATE Songs
                                SET RowVersion = randomblob(8)
                                WHERE rowid = NEW.rowid;
                            END
                        ";
                        context.Database.ExecuteSqlRaw(sqlCmd);

                        sqlCmd = @"
                                DROP VIEW IF EXISTS [PerformanceReports];
                                CREATE VIEW PerformanceReports as
                                SELECT m.ID, m.FirstName AS First_Name, m.MiddleName AS Middle_Name, m.LastName AS Last_Name,
                                    COUNT(*) AS Total_Number_Of_Performances, AVG(p.FeePaid) AS Average_Fee, MAX(p.FeePaid) AS Highest_Fee,
                                    MIN(p.FeePaid) AS Lowest_Fee
                                FROM Musicians m
                                JOIN Performances p ON m.ID = p.MusicianID
                                GROUP BY m.ID, m.FirstName, m.MiddleName, m.LastName;
                            ";
                        context.Database.ExecuteSqlRaw(sqlCmd);
                    }
                    else //The database is already created
                    {
                        if (UseMigrations)
                        {
                            context.Database.Migrate(); //Apply all migrations
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.GetBaseException().Message);
                }
                #endregion

                //Seed data needed for production and during development
                #region Seed Required Data
                try
                {
                    //Prepare some string arrays for building objects
                    string[] instruments = new string[] { "Lead Guitar", "Base Guitar", "Drums", "Keyboards", "Lead Vocals", "Backup Vocals", "Harmonica", "Violin" };
                    string[] genres = new string[] { "Classical", "Rock", "Pop", "Jazz", "Country", "Ambient", "Techno" };

                    //Instrument
                    if (!context.Instruments.Any())
                    {
                        //loop through the array of Instrument names
                        foreach (string iname in instruments)
                        {
                            Instrument inst = new Instrument()
                            {
                                Name = iname
                            };
                            context.Instruments.Add(inst);
                        }
                        context.SaveChanges();
                    }
                    //Genre
                    if (!context.Genres.Any())
                    {
                        //loop through the array of Genre names
                        foreach (string g in genres)
                        {
                            Genre genre = new()
                            {
                                Name = g
                            };
                            context.Genres.Add(genre);
                        }
                        context.SaveChanges();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.GetBaseException().Message);
                }
                #endregion

                //Seed meaningless data as sample data during development
                #region Seed Sample Data 
                if (SeedSampleData)
                {
                    try
                    {
                        //To randomly generate data
                        Random random = new Random();

                        //Prepare some string arrays for building objects
                        string[] firstNames = new string[] { "Fred", "Barney", "Wilma", "Betty", "Garrett", "Tim", "Elton", "Paul", "Shania", "Bruce" };
                        string[] lastsNames = new string[] { "Smith", "Jones", "Bloggs", "Flintstone", "Rubble", "Brown", "John", "McCartney", "Twain", "Cockburn" };
                        //Create 5 notes from Bacon ipsum
                        string[] baconNotes = new string[] { "Bacon ipsum dolor amet meatball corned beef kevin, alcatra kielbasa biltong drumstick strip steak spare ribs swine. Pastrami shank swine leberkas bresaola, prosciutto frankfurter porchetta ham hock short ribs short loin andouille alcatra. Andouille shank meatball pig venison shankle ground round sausage kielbasa. Chicken pig meatloaf fatback leberkas venison tri-tip burgdoggen tail chuck sausage kevin shank biltong brisket.", "Sirloin shank t-bone capicola strip steak salami, hamburger kielbasa burgdoggen jerky swine andouille rump picanha. Sirloin porchetta ribeye fatback, meatball leberkas swine pancetta beef shoulder pastrami capicola salami chicken. Bacon cow corned beef pastrami venison biltong frankfurter short ribs chicken beef. Burgdoggen shank pig, ground round brisket tail beef ribs turkey spare ribs tenderloin shankle ham rump. Doner alcatra pork chop leberkas spare ribs hamburger t-bone. Boudin filet mignon bacon andouille, shankle pork t-bone landjaeger. Rump pork loin bresaola prosciutto pancetta venison, cow flank sirloin sausage.", "Porchetta pork belly swine filet mignon jowl turducken salami boudin pastrami jerky spare ribs short ribs sausage andouille. Turducken flank ribeye boudin corned beef burgdoggen. Prosciutto pancetta sirloin rump shankle ball tip filet mignon corned beef frankfurter biltong drumstick chicken swine bacon shank. Buffalo kevin andouille porchetta short ribs cow, ham hock pork belly drumstick pastrami capicola picanha venison.", "Picanha andouille salami, porchetta beef ribs t-bone drumstick. Frankfurter tail landjaeger, shank kevin pig drumstick beef bresaola cow. Corned beef pork belly tri-tip, ham drumstick hamburger swine spare ribs short loin cupim flank tongue beef filet mignon cow. Ham hock chicken turducken doner brisket. Strip steak cow beef, kielbasa leberkas swine tongue bacon burgdoggen beef ribs pork chop tenderloin.", "Kielbasa porchetta shoulder boudin, pork strip steak brisket prosciutto t-bone tail. Doner pork loin pork ribeye, drumstick brisket biltong boudin burgdoggen t-bone frankfurter. Flank burgdoggen doner, boudin porchetta andouille landjaeger ham hock capicola pork chop bacon. Landjaeger turducken ribeye leberkas pork loin corned beef. Corned beef turducken landjaeger pig bresaola t-bone bacon andouille meatball beef ribs doner. T-bone fatback cupim chuck beef ribs shank tail strip steak bacon." };

                        //Create a collection of the primary keys of the Instruments
                        int[] instrumentIDs = context.Instruments.Select(a => a.ID).ToArray();
                        int instrumentIDCount = instrumentIDs.Length;

                        //Musician
                        if (!context.Musicians.Any())
                        {
                            // Start birthdate for randomly produced employees 
                            // We will subtract a random number of days from today
                            DateTime startDOB = DateTime.Today;

                            //Double loop through the arrays of names 
                            //and build the Musician as we go
                            foreach (string f in firstNames)
                            {
                                foreach (string l in lastsNames)
                                {
                                    Musician musician = new Musician()
                                    {
                                        FirstName = f,
                                        MiddleName = f.Substring(1, 1).ToUpper(),//take second letter of first name
                                        LastName = l,
                                        SIN = random.Next(113214131, 789898989).ToString(),//Big enough int for required digits
                                        Phone = random.Next(200, 999).ToString() + random.Next(2000000, 9999999).ToString(),
                                        InstrumentID = instrumentIDs[random.Next(instrumentIDCount)],
                                        DOB = startDOB.AddDays(-random.Next(6500, 25000))
                                    };
                                    context.Musicians.Add(musician);
                                    try
                                    {
                                        //Could be a duplicate SIN
                                        context.SaveChanges();
                                    }
                                    catch (Exception)
                                    {
                                        //Failed so remove it and go on to the next.
                                        //If you don't remove it from the context it
                                        //will keep trying to save it each time you 
                                        //call .SaveChanges() the the save process will stop
                                        //and prevent any other records in the que from getting saved.
                                        context.Musicians.Remove(musician);
                                    }
                                }
                            }
                        }
                        //Create a collection of the primary keys of the Musicians
                        int[] musicianIDs = context.Musicians.Select(a => a.ID).ToArray();
                        int musicianIDCount = musicianIDs.Length;

                        //Play
                        //Add a few instruments to each musician
                        if (!context.Plays.Any())
                        {
                            //i loops through the primary keys of the musicians
                            //j is just a counter so we add a few instruments to a musician
                            //k lets us step through all instruments so we can make sure each gets used
                            int k = 0;//Start with the first instrument
                            foreach (int i in musicianIDs)
                            {
                                int howMany = random.Next(1, instrumentIDCount / 2);//add a few instruments to a musician
                                for (int j = 1; j <= howMany; j++)
                                {
                                    k = (k >= instrumentIDCount) ? 0 : k;//Resets counter k to 0 if we have run out of instruments
                                    Play p = new Play()
                                    {
                                        MusicianID = i,
                                        InstrumentID = instrumentIDs[k]
                                    };
                                    context.Plays.Add(p);
                                    k++;
                                }
                            }
                            context.SaveChanges();
                        }

                        //Create a collection of the primary keys of the Genres
                        int[] genreIDs = context.Genres.Select(g => g.ID).ToArray();
                        int genreIDCount = genreIDs.Length;
                        //Album
                        if (!context.Albums.Any())
                        {
                            context.Albums.AddRange(
                             new Album
                             {
                                 Name = "Rocket Food",
                                 YearProduced = "2000",
                                 Price = 19.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "Songs of the Sea",
                                 YearProduced = "1999",
                                 Price = 9.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "The Horse",
                                 YearProduced = "1929",
                                 Price = 99.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "Music From Away",
                                 YearProduced = "1999",
                                 Price = 9.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "Life",
                                 YearProduced = "1988",
                                 Price = 19.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "Small Minds, Big Hearts",
                                 YearProduced = "1967",
                                 Price = 12.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "The Cow",
                                 YearProduced = "2010",
                                 Price = 21.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             },
                             new Album
                             {
                                 Name = "Freedom",
                                 YearProduced = "2012",
                                 Price = 29.99d,
                                 GenreID = genreIDs[random.Next(genreIDs.Count())]
                             });
                            context.SaveChanges();
                        }

                        //Create a collection of the primary keys of the Albums
                        int[] albumIDs = context.Albums.Select(a => a.ID).ToArray();
                        int albumIDCount = albumIDs.Length;

                        //Song
                        if (!context.Songs.Any())
                        {
                            // Start date for random recording dates
                            // We will subtract a random number of days from today
                            DateTime startDate = DateTime.Today;
                            int counter = 1; //Used to give seperate Genres to some Songs

                            //Double loop through the arrays of names 
                            //and build song title as you go
                            foreach (string l in lastsNames)
                            {
                                foreach (string f in firstNames)
                                {
                                    string name1 = l.Substring(1);
                                    name1 = char.ToUpper(name1[0]) + name1.Substring(1);
                                    string name2 = f.Substring(1);
                                    name2 = char.ToUpper(name2[0]) + name2.Substring(1);
                                    Song s = new()
                                    {
                                        Title = name1 + " " + name2,//looks silly but gives unique names for the songs,
                                        DateRecorded = startDate.AddDays(-random.Next(30, 1000)),
                                        //GenreID = genreIDs[random.Next(genreIDCount)],
                                        AlbumID = albumIDs[random.Next(albumIDCount)]
                                    };
                                    if (counter % 5 == 0)//Every fifth Song gets a Genre
                                    {
                                        s.GenreID = genreIDs[random.Next(genreIDCount)];
                                    }
                                    counter++;
                                    context.Songs.Add(s);
                                    try
                                    {
                                        //Could be a duplicate 
                                        context.SaveChanges();
                                    }
                                    catch (Exception)
                                    {
                                        context.Songs.Remove(s);
                                    }
                                }
                            }
                            context.SaveChanges();
                        }
                        //Create a collection of the primary keys of the Songss
                        int[] songIDs = context.Songs.Select(a => a.ID).ToArray();
                        int songIDCount = songIDs.Length;

                        //Performance
                        //Add a few musicians as performers on each song
                        if (!context.Performances.Any())
                        {
                            //i loops through the primary keys of the songs
                            //j is just a counter so we add a few musicians to a song
                            //k lets us step through all musicians so we can make sure each gets used
                            int k = 0;//Start with the first Musician
                            foreach (int i in songIDs)
                            {
                                int howMany = random.Next(1, 7);//How many musicians on a song
                                howMany = (howMany > musicianIDCount) ? musicianIDCount - 1 : howMany; //Don't try to assign more musicians then are in the system
                                for (int j = 1; j <= howMany; j++)
                                {
                                    k = (k >= musicianIDCount) ? 0 : k;
                                    Performance p = new()
                                    {
                                        Comments = baconNotes[random.Next(5)],
                                        FeePaid = random.Next(500),
                                        SongID = i,
                                        MusicianID = musicianIDs[k],
                                        InstrumentID = context.Musicians.Find(musicianIDs[k]).InstrumentID//Get the primary instrument of the musician
                                    };
                                    context.Performances.Add(p);
                                    k++;
                                }
                            }
                            context.SaveChanges();
                        }


                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine(ex.GetBaseException().Message);
                    }
                }

                #endregion

            }
        }
    }
}