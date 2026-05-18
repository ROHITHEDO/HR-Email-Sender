using OfficeOpenXml;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

class Program
{
    static async Task Main()
    {
        ExcelPackage.License.SetNonCommercialPersonal("Rohith");

        var file = new FileInfo("hr_list.xlsx");

        using var package = new ExcelPackage(file);

        var sheet = package.Workbook.Worksheets["Persistent"];

        int row = 1;

        while (sheet.Cells[row, 1].Value != null)
        {
            string email = sheet.Cells[row, 1].Text;

            try
            {
                Console.WriteLine($"Sending to {email}");

                var message = new MimeMessage();

                message.From.Add(
                    new MailboxAddress(
                        "Rohith",
                        "rohith.edo@gmail.com"));

                message.To.Add(
                    MailboxAddress.Parse(email));

                message.Subject =
                    "Application for .NET Developer Role";

                var builder = new BodyBuilder();

                builder.TextBody = @"
Hi,

I came across your requirement for a .NET Developer with 4.5+ years of experience, and I genuinely feel this role aligns perfectly with what I’ve been doing and what I’m passionate about.

Skills- 

Backend: .NET Core, Web API, Entity Framework (Code First & DB First), LINQ, SQL Server

Frontend: Angular, JavaScript, HTML, CSS

DevOps/Tools: Git, Azure DevOps, Postman, Swagger, IIS

Database: SQL Server, Cosmos DB, Azure storages

Other: Strong debugging skills, clean architecture, role-based access, RESTful APIs, stored procedures, Dapper

I’ve worked directly with clients, understood requirements, and delivered modules end-to-end and I’m looking for a company where I can grow faster, work with talented people, and contribute meaningfully.

Please let me know if we can connect or schedule an interview any time soon.

Thanks & Regards,
Rohith
Software Engineer 
https://www.linkedin.com/in/rohith-developer
";

                builder.Attachments.Add("RohithKumar_MAY_2026.pdf");

                message.Body = builder.ToMessageBody();

                using var client =
                    new MailKit.Net.Smtp.SmtpClient();

                await client.ConnectAsync(
                    "smtp.gmail.com",
                    587,
                    SecureSocketOptions.StartTls);

                await client.AuthenticateAsync(
                    "rohith.edo@gmail.com",
                    "satwshxqpczpcnyq");

                await client.SendAsync(message);

                await client.DisconnectAsync(true);

                sheet.Cells[row, 2].Value = "Sent";

                Console.WriteLine("SUCCESS");
            }
            catch (Exception ex)
            {
                sheet.Cells[row, 2].Value = "Failed";
                sheet.Cells[row, 3].Value = ex.Message;

                Console.WriteLine(ex.Message);
            }

            row++;

            await Task.Delay(3000);
        }

        package.Save();

        Console.WriteLine("Done");
    }
}