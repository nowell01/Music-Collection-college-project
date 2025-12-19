using System.ComponentModel.DataAnnotations;

namespace MVC_Music.ViewModels
{
    public class PerformanceReportVM
    {
        public int ID { get; set; }

        [Display(Name = "Musician")]
        public string Summary
        {
            get
            {
                return First_Name
                    + (string.IsNullOrEmpty(Middle_Name) ? " " :
                        (" " + (char?)Middle_Name[0] + ". ").ToUpper())
                    + Last_Name;
            }
        }

        public string First_Name { get; set; } = "";

        public string? Middle_Name { get; set; }

        public string Last_Name { get; set; } = "";

        [Display(Name = "Average Fee")]
        [DataType(DataType.Currency)]
        public double? Average_Fee { get; set; }

        [Display(Name = "Highest Fee")]
        [DataType(DataType.Currency)]
        public double? Highest_Fee { get; set; }

        [Display(Name = "Lowest Fee")]
        [DataType(DataType.Currency)]
        public double? Lowest_Fee { get; set; }

        [Display(Name = "Total Number of Performances")]
        public int? Total_Number_Of_Performances { get; set; }
    }
}
