using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreateClosureData
{
    class Ticket
    {
        public string Description { get; set; }
        public string SolvedBy { get; set; }
        public string Requestor { get; set; }
        public string Category { get; set; }

        public Ticket()
        {

        }
        public Ticket(string description, string solved, string requestor, string category)
        {
            this.Description = description;
            this.SolvedBy = solved;
            this.Requestor = requestor;
            this.Category = category;
        }
    }
}
