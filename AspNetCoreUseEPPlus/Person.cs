using OfficeOpenXml.Attributes;
using System.ComponentModel;

namespace AspNetCoreUseEPPlus
{
    public class Person
    {
        public Person(string firstName, string lastName, int height, DateTime birthDate)
        {
            FirstName = firstName;
            LastName = lastName;
            Height = height;
            BirthDate = birthDate;
        }

        [DisplayName("FirstName")]
        public string FirstName { get; set; }

        [DisplayName("LastName")]
        public string LastName { get; set; }

        [DisplayName("身高")]
        public int Height { get; set; }

        [EpplusTableColumn(Header = "生日")]
        public DateTime BirthDate { get; set; }

        /// <summary>
        /// 忽略
        /// </summary>
        [EpplusIgnore]
        public string Password { get; set; }
    }
}
