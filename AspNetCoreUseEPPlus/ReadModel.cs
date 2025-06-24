using System.ComponentModel;

namespace AspNetCoreUseEPPlus
{
    public class ReadModel
    {
        [DisplayName("Name")]
        public string Name { get; set; }

        [DisplayName("Age")]
        public int Age { get; set; }

        [DisplayName("Address")]
        public string Address { get; set; }
    }
}
