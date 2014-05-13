using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace elastic_outlook_indexer
{
    public class Recipient
    {
        public string Address { get; set; }
        public AddressEntry AddressEntry { get; set; }
        public string EntryId { get; set; }
        public int Index { get; set; }
        public string Name { get; set; }
        public int Type { get; set; }
    }
}
