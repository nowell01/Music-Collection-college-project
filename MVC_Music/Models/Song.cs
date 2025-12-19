using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace MVC_Music.Models
{
    public class Song : Auditable
    {
        public int ID { get; set; }

        [Display(Name = "Song")]
        public string Summary
        {
            get
            {
                return (Genre == null) ? Title : Title + " (" + Genre?.Name + ")";
            }
        }

        [Required(ErrorMessage = "You cannot leave the Song title blank.")]
        [StringLength(80, ErrorMessage = "Song title cannot be more than 80 characters long.")]
        public string Title { get; set; } = "";

        [Display(Name = "Date Recorded")]
        [Required(ErrorMessage = "You must enter the Date Recorded")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime DateRecorded { get; set; }

        [ScaffoldColumn(false)]
        [Timestamp]
        public Byte[]? RowVersion { get; set; }//Added for concurrency

        [Display(Name = "Album")]
        [Required(ErrorMessage = "You must select the Album.")]
        public int AlbumID { get; set; }
        public Album? Album { get; set; }

        [Display(Name = "Genre")]
        public int? GenreID { get; set; }
        public Genre? Genre { get; set; }

        public ICollection<Performance> Performances { get; set; } = new HashSet<Performance>();
    }
}
