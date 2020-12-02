using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NLog;
using System.IO;
using ImportData.Logic;
using System.Xml.Linq;
using Sungero.Domain.Client;

using Sungero.Domain.Client;
using Sungero.Domain.ClientLinqExpressions;

namespace ImportData
{
  public class CompaniesIMVO
	{
		/// <summary>
		/// Чтение значения из атрибута XML
		/// </summary>
		/// <param name="attribute"></param>
		/// <returns>string</returns>
		private static string GetXMLValue(XAttribute attribute)
		{
			if (attribute == null)
				return "";
			else
				return attribute.Value.ToString();
		}


		public static void Procces(string xlsxPath, Logger logger)
    {		
			int countOrgsInXML = 0;
			int countOrgsToImport = 0;
			int countImportedOrgs = 0;
			int countCreatedRegions = 0;
			int countCreatedCities = 0;
			int countCreatedContacts = 0;

			logger.Info("===================Чтение строк из файла===================");
      var watch = System.Diagnostics.Stopwatch.StartNew();

			XDocument xdoc = XDocument.Load(xlsxPath);

			List<CompanyFromXML> companies = new List<CompanyFromXML>();
			//List

			// Парсим XML
			foreach (XElement org in xdoc.Root.Element("DECLARBODY").Elements("ORG"))
			{
				countOrgsInXML++;
				bool isCompanyCorrect = true;
				CompanyFromXML company = new CompanyFromXML();

				// Наименование
				XAttribute nameAttribute = org.Attribute("NAMES");
				if (!string.IsNullOrWhiteSpace(nameAttribute.Value))
					company.Name = nameAttribute.Value;
				else
					isCompanyCorrect = false;

				// ЕДРПОУ
				XAttribute egrpouAttribute = org.Attribute("ZKPO");
				if (!string.IsNullOrWhiteSpace(egrpouAttribute.Value))
					company.EGRPOU = egrpouAttribute.Value;
				else
					isCompanyCorrect = false;

				if (isCompanyCorrect)
				{					
					// Юридическое наименование
					XAttribute legalNameAttribute = org.Attribute("NAME");
					if (!string.IsNullOrWhiteSpace(legalNameAttribute.Value))
						company.LegalName = legalNameAttribute.Value;

					// ИНН
					XAttribute innAttribute = org.Attribute("PN");
					if (!string.IsNullOrWhiteSpace(innAttribute.Value))
						company.INN = innAttribute.Value;

					// Код
					XAttribute kodAttribute = org.Attribute("NP");
					if (!string.IsNullOrWhiteSpace(kodAttribute.Value))
						company.Kod = kodAttribute.Value;

					// Телефон
					XAttribute telAttribute = org.Attribute("TEL");
					if (!string.IsNullOrWhiteSpace(telAttribute.Value))
						company.Phone = telAttribute.Value;

					// Почта
					XAttribute emailAttribute = org.Attribute("EMAIL");
					if (!string.IsNullOrWhiteSpace(emailAttribute.Value))
						company.Email = emailAttribute.Value;

					// Сайт
					XAttribute siteAttribute = org.Attribute("WEB");
					if (!string.IsNullOrWhiteSpace(siteAttribute.Value))
						company.Web = siteAttribute.Value;

					// Адрес
					foreach (XElement adr in org.Elements("ADRES"))
					{
						XAttribute typeadresAttribute = adr.Attribute("ADREST");
						if (typeadresAttribute != null)
						{
							XAttribute indexAttribute = adr.Attribute("PINDEX");
							XAttribute cityAttribute = adr.Attribute("CITY");
							XAttribute typenamexAttribute = adr.Attribute("VULT_NAME");
							XAttribute vulnameAttribute = adr.Attribute("VUL_NAME");
							XAttribute officetAttribute = adr.Attribute("OFFICET_NAME");
							XAttribute budAttribute = adr.Attribute("BUD");
							XAttribute officeAttribute = adr.Attribute("OFFICE");

							string address = string.Format("{0} {1} {2} {3} {4} {5} {6}", GetXMLValue(indexAttribute), GetXMLValue(cityAttribute), GetXMLValue(typenamexAttribute), GetXMLValue(vulnameAttribute), GetXMLValue(officetAttribute), GetXMLValue(budAttribute), GetXMLValue(officeAttribute));

							switch (typeadresAttribute.Value)
							{
								case "":
								case "Юридична":
									company.LegalAdress = address ?? "";
									// Регион
									XAttribute regionAttribute = adr.Attribute("REGION_NAME");
									if (!string.IsNullOrWhiteSpace(regionAttribute.Value))
										company.RegionName = regionAttribute.Value;

									// Населенный пункт										
									if (!string.IsNullOrWhiteSpace(cityAttribute.Value))
										company.CytyName = cityAttribute.Value;

									break;
								case "Поштова":
									company.PostAdress = address ?? "";
									break;
								default:
									break;
							}
						}
					}

					// Банк
					foreach (XElement acc in org.Elements("ACC"))
					{
						XAttribute osnAttribute = acc.Attribute("OSN");
						if (osnAttribute.Value == "1")
						{
							// IBAN
							XAttribute accAttribute = acc.Attribute("ACC");
							if (!string.IsNullOrWhiteSpace(accAttribute.Value))
								company.IBAN = accAttribute.Value;

							// MFO
							XAttribute mfoAttribute = acc.Attribute("MFO");
							if (!string.IsNullOrWhiteSpace(mfoAttribute.Value))
								company.BankMFO = mfoAttribute.Value;
						}
					}

					// Контакты
					List<ContactFromXML> contacts = new List<ContactFromXML>();
					foreach (XElement cont in org.Elements("CONTACT"))
					{
						XAttribute fioAttribute = cont.Attribute("NAME");
						XAttribute postAttribute = cont.Attribute("POSADA");
						XAttribute conttelAttribute = cont.Attribute("TEL");
						XAttribute contmailAttribute = cont.Attribute("EMAIL");
						if (!string.IsNullOrWhiteSpace(fioAttribute.Value))
						{
							ContactFromXML contact = new ContactFromXML();
							contact.FIO = GetXMLValue(fioAttribute);
							contact.Position = GetXMLValue(postAttribute);
							contact.Phone = GetXMLValue(conttelAttribute);
							contact.Email = GetXMLValue(contmailAttribute);
							contacts.Add(contact);
						}
					}
					if (contacts.Any())
						company.Contacts = contacts;

					companies.Add(company);
				}				
			}
			
			watch.Stop();
      var elapsedMs = watch.ElapsedMilliseconds;
      logger.Info($"Времени затрачено на чтение строк из файла: {elapsedMs} мс");

			logger.Info("======================Импорт сущностей=====================");

			if (companies.Any())
			{
				countOrgsToImport = companies.Count();
				watch.Restart();
				var isActive = Sungero.CoreEntities.DatabookEntry.Status.Active;

				// Создать страну Украина, если таковой нет
				Sungero.Commons.ICountry countryUA = null;
				using (var session = new Session())
				{					
					var countries = Enumerable.ToList(session.GetAll<Sungero.Commons.ICountry>().Where(c => c.Code == "980" || c.Name == "Украина" || c.Name == "Україна"));
					countryUA = (Enumerable.FirstOrDefault<Sungero.Commons.ICountry>(countries));										
					if (countryUA == null)
					{
						countryUA = session.Create<Sungero.Commons.ICountry>();						
						countryUA.Name = "Україна";
						countryUA.Code = "980";
						countryUA.Status = isActive;
						countryUA.Save();
						session.SubmitChanges();
						logger.Info("Создано страну Україна");
					}					
				}

				int i = 1;
				foreach (CompanyFromXML companyInfo in companies)
				{					
					logger.Info($"Обработка организации №{i.ToString()} Наименование:{companyInfo.Name} ЕДРПОУ: {companyInfo.EGRPOU}");
					try
					{
						using (var session = new Session())
						{							
							var companiesUA = Enumerable.ToList(session.GetAll<litiko.UAadditions.ICompany>().Where(c => c.EGRPOUlitiko.Equals(companyInfo.EGRPOU)));
							var company = (Enumerable.FirstOrDefault<litiko.UAadditions.ICompany>(companiesUA));
							if (company != null)
								logger.Info($"Организация уже существует, ЕГРПОУ:{companyInfo.EGRPOU}. Импорт не требуется.");
							else
							{
								company = session.Create<litiko.UAadditions.ICompany>();								
								company.Name = companyInfo.Name;
								company.LegalName = companyInfo.LegalName;
								company.EGRPOUlitiko = companyInfo.EGRPOU;
								company.TIN = companyInfo.INN;
								company.Code = companyInfo.Kod;
								company.Status = isActive;
								company.PostalAddress = companyInfo.PostAdress;
								company.LegalAddress = companyInfo.LegalAdress;
								company.Phones = companyInfo.Phone;
								company.Email = companyInfo.Email;
								company.Homepage = companyInfo.Web;
								company.Account = companyInfo.IBAN;
								company.Note = "[Загружено автоматически]";

								//Регион
								Sungero.Commons.IRegion region = null;
								if (!string.IsNullOrWhiteSpace(companyInfo.RegionName))
								{
									var regions = Enumerable.ToList(session.GetAll<Sungero.Commons.IRegion>().Where(r => r.Name.Equals(companyInfo.RegionName) && r.Country.Equals(countryUA)));
									region = (Enumerable.FirstOrDefault<Sungero.Commons.IRegion>(regions));									
									if (region != null)
										company.Region = region;
									else
									{
										// Создаем новый регион
										//region = Sungero.Commons.Regions.Create();
										region = session.Create<Sungero.Commons.IRegion>();
										region.Country = countryUA;
										region.Name = companyInfo.RegionName;
										region.Status = isActive;
										region.Save();

										company.Region = region;
										logger.Info($"Создан регион: {companyInfo.RegionName}");
										countCreatedRegions++;
									}
								}

								// Город - ПРОВЕРИТЬ, СОЗДАДИТСЯ ЛИ РЕГИОН И ГОРОД В РАМКАХ ОДНОЙ СЕССИИ !!!
								Sungero.Commons.ICity city = null;
								if (!string.IsNullOrWhiteSpace(companyInfo.CytyName) && region != null)
								{
									var cities = Enumerable.ToList(session.GetAll<Sungero.Commons.ICity>().Where(c => c.Country.Equals(countryUA) && c.Region.Equals(region) && c.Name.Equals(companyInfo.CytyName)));
									city = (Enumerable.FirstOrDefault<Sungero.Commons.ICity>(cities));
									if (city != null)
										company.City = city;
									else
									{
										// Создаем населенный пункт										
										city = session.Create<Sungero.Commons.ICity>();
										city.Country = countryUA;
										city.Region = region;
										city.Name = companyInfo.CytyName;
										city.Status = isActive;
										city.Save();

										company.City = city;
										logger.Info($"Создан населенный пункт: {companyInfo.CytyName}");
										countCreatedCities++;
									}
								}

								// Банк
								litiko.UAadditions.IBank bank = null;
								if (!string.IsNullOrWhiteSpace(companyInfo.BankMFO))
								{
									var banks = Enumerable.ToList(session.GetAll<litiko.UAadditions.IBank>().Where(b => b.MFOlitiko.Equals(companyInfo.BankMFO)));
									bank = (Enumerable.FirstOrDefault<litiko.UAadditions.IBank>(banks));
									if (bank != null)
										company.Bank = bank;
								}

								company.Save();

								// Контакты
								if (companyInfo.Contacts != null)
								{
									Sungero.Parties.IContact contact = null;
									foreach (var contactInfo in companyInfo.Contacts)
									{
										var contacts = Enumerable.ToList(session.GetAll<Sungero.Parties.IContact>().Where(c => c.Company.Equals(company) && c.Name.Equals(contactInfo.FIO)));
										contact = (Enumerable.FirstOrDefault<Sungero.Parties.IContact>(contacts));
										if (contact != null)
											logger.Info($"Контакт уже существует:{contactInfo.FIO}. Импорт не требуется.");
										else
										{
											contact = session.Create<Sungero.Parties.IContact>();
											contact.Name = contactInfo.FIO;
											contact.Company = company;
											contact.JobTitle = contactInfo.Position;
											contact.Phone = contactInfo.Phone;
											contact.Email = contactInfo.Email;
											contact.Save();
											logger.Info($"Создан контакт: {contactInfo.FIO}");
											countCreatedContacts++;
										}										
									}
								}

								session.SubmitChanges();
								countImportedOrgs++;
								logger.Info("Организация успешно сохранена.");
							}							
						}
					}
					catch (Exception ex)
					{
						logger.Error("Ошибка при импорте: " + ex.Message);
					}
					i++;
				}

				watch.Stop();
				elapsedMs = watch.ElapsedMilliseconds;
				logger.Info($"Времени затрачено на импорт сущностей: {elapsedMs} мс");
			}

			logger.Info("======================Результаты=====================");
			logger.Info($"Организаций в исходном xml-файле: {countOrgsInXML}");
			logger.Info($"Извлечено организаций, подходящих для импорта (есть имя и ЕДРПОУ): {countOrgsToImport}");
			var percent = (double)(countImportedOrgs) / (double)(countOrgsToImport) * 100.00;
			logger.Info($"Импортировано организаций: {countImportedOrgs} шт. или {percent}%");
			logger.Info($"Создано регионов: {countCreatedRegions} шт.");
			logger.Info($"Создано населенных пунктов: {countCreatedCities} шт.");
			logger.Info($"Создано контактов: {countCreatedContacts} шт.");			
		}
	}
}
