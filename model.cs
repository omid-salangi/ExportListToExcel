using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    public class IExcelReport
    {
        public string Id { get; set; }
        public string Username { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Password { get; set; }
        public DateTime CreatedOn { get; set; }

        public static List<IExcelReport> SeedData()
        {
            return new List<IExcelReport>
            {
                new IExcelReport
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 1",
                    LastName = "Test 1",
                    Username = "TestUsername1",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },
                new IExcelReport
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 2",
                    LastName = "Test 2",
                    Username = "TestUsername2",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },
                new IExcelReport
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 3",
                    LastName = "Test 3",
                    Username = "TestUsername3",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                },new IExcelReport
                {
                    Id = Guid.NewGuid().ToString(),
                    FirstName = "Test 4",
                    LastName = "Test 4",
                    Username = "TestUsername4",
                    Password = "123",
                    CreatedOn = DateTime.UtcNow
                }
            };
        }
    }
}
