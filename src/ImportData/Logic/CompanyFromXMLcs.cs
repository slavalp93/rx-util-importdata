using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportData.Logic
{
	public class CompanyFromXML
	{
		public string Name { get; set; }
		public string LegalName { get; set; }
		public string EGRPOU { get; set; }
		public string INN { get; set; }
		public string Kod { get; set; }

		public string RegionName { get; set; }
		public string CytyName { get; set; }
		public string PostAdress { get; set; }
		public string LegalAdress { get; set; }
		public string Phone { get; set; }
		public string Email { get; set; }
		public string Web { get; set; }

		public string IBAN { get; set; }
		public string BankMFO { get; set; }

		public List<ContactFromXML> Contacts { get; set; }
	}

	public class ContactFromXML
	{
		public string FIO { get; set; }
		public string Position { get; set; }
		public string Phone { get; set; }
		public string Email { get; set; }
	}
}
