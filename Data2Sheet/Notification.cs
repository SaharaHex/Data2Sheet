using MySql.Data.MySqlClient;
using System;

namespace Data2Sheet
{
    /// <summary>
    /// Create Notification for the application.
    /// </summary>
    internal class Notification
    {
        public static void MK_LowStockLaptop(MySqlDataReader reader, int quantityTrigger)
        {
            string totalCount = reader.GetString(0);
            int.TryParse(totalCount, out int i);
            string message = "Low Stock Laptop: " + totalCount; 

            if (i <= quantityTrigger)
            {
                Email em = new Email("test100@test.com", message, "Subject Line", "");
                em.SendMail();
            }
            Console.WriteLine(totalCount);
        }
    }
}
