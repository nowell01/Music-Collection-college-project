using System.ComponentModel.DataAnnotations;

namespace MVC_Music.Models
{
    public class Musician : Auditable, IValidatableObject
    {
        public int ID { get; set; }

        [Display(Name = "Musician")]
        public string Summary
        {
            get
            {
                return FirstName
                    + (string.IsNullOrEmpty(MiddleName) ? " " :
                        (" " + (char?)MiddleName[0] + ". ").ToUpper())
                    + LastName;
            }
        }

        [Display(Name = "Musician")]
        public string FormalName
        {
            get
            {
                return LastName + ", " + FirstName
                    + (string.IsNullOrEmpty(MiddleName) ? "" :
                        (" " + (char?)MiddleName[0] + ".").ToUpper());
            }
        }

        public int Age
        {
            get
            {
                DateTime today = DateTime.Today;
                int a = today.Year - DOB.Year
                    - ((today.Month < DOB.Month || (today.Month == DOB.Month && today.Day < DOB.Day) ? 1 : 0));
                return a; /*Note: You could add .PadLeft(3) but spaces disappear in a web page. */
            }
        }

        [Display(Name = "Age (DOB)")]
        public string AgeSummary => Age + " (" + DOB.ToString("yyyy-MM-dd") + ")";

        [Display(Name = "SIN")]
        public string SINFormatted
        {
            get
            {
                return SIN.Substring(0, 3) + "-" + SIN.Substring(3, 3) + "-" + SIN.Substring(6, 3);
            }
        }

        [Display(Name = "Phone")]
        public string PhoneFormatted
        {
            get
            {
                return "(" + Phone.Substring(0, 3) + ") " + Phone.Substring(3, 3) + "-" + Phone[6..];
            }
        }

        [Display(Name = "First Name")]
        [Required(ErrorMessage = "You cannot leave the first name blank.")]
        [StringLength(30, ErrorMessage = "First name cannot be more than 30 characters long.")]
        public string FirstName { get; set; } = "";

        [Display(Name = "Middle Name")]
        [StringLength(30, ErrorMessage = "Middle name cannot be more than 30 characters long.")]
        public string? MiddleName { get; set; } = "";

        [Display(Name = "Last Name")]
        [Required(ErrorMessage = "You cannot leave the last name blank.")]
        [StringLength(50, ErrorMessage = "Last name cannot be more than 50 characters long.")]
        public string LastName { get; set; } = "";

        [Required(ErrorMessage = "Phone number is required.")]
        [RegularExpression(@"^[2-9]\d{2}[2-9]\d{6}$", ErrorMessage = "Enter a valid 10-digit phone number.")]
        [DataType(DataType.PhoneNumber)]
        [StringLength(10)]
        public string Phone { get; set; } = "";

        [Required(ErrorMessage = "You must enter the Date of Birth")]
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:yyyy-MM-dd}", ApplyFormatInEditMode = true)]
        public DateTime DOB { get; set; }

        [Required(ErrorMessage = "You cannot leave the SIN blank.")]
        //First digit cannot be a zero but is followed by any 8 digits.  It is true that
        //SIN numbers starting with 8 or 9 are only temporary ones but we will ignor that.
        //We will also not worry about validating the SIN with the Luhn Algorithm.
        [RegularExpression(@"^[1-9]\d{8}$", ErrorMessage = "Please enter a valid 9-digit SIN.")]
        [StringLength(9)]//DS Note: we only include this to limit the size of the database field
        public string SIN { get; set; } = "";

        public MusicianPhoto? MusicianPhoto { get; set; }
        public MusicianThumbnail? MusicianThumbnail { get; set; }

        [ScaffoldColumn(false)]
        [Timestamp]
        public Byte[]? RowVersion { get; set; }//Added for concurrency

        [Display(Name = "Primary Instrument")]
        [Required(ErrorMessage = "You must select the principal instrument the musician plays.")]
        public int InstrumentID { get; set; }

        [Display(Name = "Primary Inst.")]
        public Instrument? Instrument { get; set; }

        [Display(Name = "All Instruments")]
        public ICollection<Play> Plays { get; set; } = new HashSet<Play>();

        public ICollection<Performance> Performances { get; set; } = new HashSet<Performance>();

        [Display(Name = "Documents")]
        public ICollection<MusicianDocument> MusicianDocuments { get; set; } = new HashSet<MusicianDocument>();

        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            if (DOB > DateTime.Today)
            {
                yield return new ValidationResult("Date of Birth cannot be in the future.", new[] { "DOB" });
            }
            else if (Age < 7)
            {
                yield return new ValidationResult("Musician must be at least 7 years old.", new[] { "DOB" });
            }
        }
    }
}
