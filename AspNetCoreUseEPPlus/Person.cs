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

        [DisplayName("The persons first name")]
        public string FirstName { get; set; }

        public string LastName { get; set; }

        public int Height { get; set; }

        public DateTime BirthDate { get; set; }
    }
}
