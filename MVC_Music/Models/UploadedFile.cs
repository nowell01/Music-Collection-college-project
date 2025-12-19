using System.ComponentModel.DataAnnotations;

namespace MVC_Music.Models
{
    public class UploadedFile
    {
        public int ID { get; set; }

        [Display(Name = "File Name")]
        [StringLength(255, ErrorMessage = "The name of the file cannot be more than 255 characters.")]
        public string? FileName { get; set; }

        [Display(Name = "Type of File")]
        [StringLength(255, ErrorMessage = "The mime type of the file cannot be more than 255 characters.")]
        public string? MimeType { get; set; }

        [StringLength(2000, ErrorMessage = "File description cannot be more than 2000 characters.")]
        [DataType(DataType.MultilineText)]
        public string? Description { get; set; }

        public FileContent? FileContent { get; set; } = new FileContent();
    }
}

