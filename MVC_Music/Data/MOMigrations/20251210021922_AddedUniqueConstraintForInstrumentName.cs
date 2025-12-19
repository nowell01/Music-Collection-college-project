using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace MVC_Music.Data.MOMigrations
{
    /// <inheritdoc />
    public partial class AddedUniqueConstraintForInstrumentName : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateIndex(
                name: "IX_Instruments_Name",
                table: "Instruments",
                column: "Name",
                unique: true);
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropIndex(
                name: "IX_Instruments_Name",
                table: "Instruments");
        }
    }
}
