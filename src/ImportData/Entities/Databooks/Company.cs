using System;
using System.Collections.Generic;
using NLog;
using System.Linq;
using Sungero.Domain.Client;
using Sungero.Domain.ClientLinqExpressions;
using System.Text.RegularExpressions;

namespace ImportData
{
  class Company : Entity
  {
    public int PropertiesCount = 19;
    /// <summary>
    /// Получить наименование число запрашиваемых параметров.
    /// </summary>
    /// <returns>Число запрашиваемых параметров.</returns>
    public override int GetPropertiesCount()
    {
      return PropertiesCount;
    }

    /// <summary>
    /// Сохранение сущности в RX.
    /// </summary>
    /// <param name="shift">Сдвиг по горизонтали в XLSX документе. Необходим для обработки документов, составленных из элементов разных сущностей.</param>
    /// <param name="logger">Логировщик.</param>
    /// <returns>Число запрашиваемых параметров.</returns>
    public override IEnumerable<Structures.ExceptionsStruct> SaveToRX(NLog.Logger logger, bool supplementEntity, int shift = 0)
    {
      
      var exceptionList = new List<Structures.ExceptionsStruct>();

      using (var session = new Session())
      {
        var name = this.Parameters[shift + 0].Trim();
        if (string.IsNullOrEmpty(name))
        {
          var message = string.Format("Не заполнено поле \"Наименование\".");
          exceptionList.Add(new Structures.ExceptionsStruct {ErrorType = "Error", Message = message});
          logger.Error(message);
          return exceptionList;
        }
        var legalName = this.Parameters[shift + 1].Trim();
        var counterparty = BusinessLogic.GetConterparty(session, this.Parameters[shift + 2].Trim(), exceptionList, logger);
        if (!string.IsNullOrEmpty(this.Parameters[shift + 2].Trim()) && counterparty == null)
        {
          counterparty = BusinessLogic.CreateConterparty(session, this.Parameters[shift + 2].Trim(), exceptionList, logger);
          //var message = string.Format("Не найдена Головная организация \"{1}\". Наименование организации: \"{0}\". ", name, this.Parameters[shift + 2].Trim());
          //exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
          //logger.Warn(message);
        }
        var headCompany = Sungero.Parties.Companies.As(counterparty);

        var nonresident = this.Parameters[shift + 3] == "Да" ? true : false;
        var tin = this.Parameters[shift + 4].Trim();
        var trrc = this.Parameters[shift + 5].Trim();
        var psrn = this.Parameters[shift + 6].Trim();
        var nceo = this.Parameters[shift + 7].Trim();
        var ncea = this.Parameters[shift + 8].Trim();
        var city = BusinessLogic.GetCity(session, this.Parameters[shift + 9].Trim(), exceptionList, logger);
        if (!string.IsNullOrEmpty(this.Parameters[shift + 9].Trim()) && city == null)
        {
          var message = string.Format("Не найден Населенный пункт \"{1}\". Наименование организации: \"{0}\". ", name, this.Parameters[shift + 9].Trim());
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
          logger.Warn(message);
        }
        var region = BusinessLogic.GetRegion(session, this.Parameters[shift + 10].Trim(), exceptionList, logger);
        if (!string.IsNullOrEmpty(this.Parameters[shift + 10].Trim()) && region == null)
        {
          var message = string.Format("Не найден Регион \"{1}\". Наименование организации: \"{0}\". ", name, this.Parameters[shift + 10].Trim());
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
          logger.Warn(message);
        }
        var legalAdress = this.Parameters[shift + 11].Trim();
        var postalAdress = this.Parameters[shift + 12].Trim();
        var phones = this.Parameters[shift + 13].Trim();
        var email = this.Parameters[shift + 14].Trim();
        var homepage = this.Parameters[shift + 15].Trim();
        var note = this.Parameters[shift + 16].Trim();
        var account = this.Parameters[shift + 17].Trim();
        var bank = BusinessLogic.GetBank(session, this.Parameters[shift + 18].Trim(), exceptionList, logger);
        if (!string.IsNullOrEmpty(this.Parameters[shift + 18]) && bank == null)
        {
          var message = string.Format("Не найден Банк \"{1}\". Наименование организации: \"{0}\". ", name, this.Parameters[shift + 18].Trim());
          exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Warn, Message = message });
          logger.Warn(message);
        }
        try
        {
          // Проверка ИНН.
          var resultTIN = Sungero.Parties.PublicFunctions.Counterparty.CheckTin(tin, true);
          if (!string.IsNullOrEmpty(resultTIN))
          {
            var message = string.Format("Компания не может быть импортирована. Некорректный ИНН. Наименование: \"{0}\", ИНН: {1}. {2}", name, tin, resultTIN);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);
            return exceptionList;
          }

          // Проверка КПП.
          var resultTRRC = BusinessLogic.CheckTrrcLength(trrc);
          if (!string.IsNullOrEmpty(resultTRRC))
          {
            var message = string.Format("Компания не может быть импортирована. Некорректный КПП. Наименование: \"{0}\", КПП: {1}. {2}", name, trrc, resultTRRC);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);
            return exceptionList;
          }

          // Проверка ОГРН.
          var resultPSRN = BusinessLogic.CheckPsrnLength(psrn);
          if (!string.IsNullOrEmpty(resultPSRN))
          {
            var message = string.Format("Компания не может быть импортирована. Некорректный ОГРН. Наименование: \"{0}\", ОГРН: {1}. {2}", name, psrn, resultPSRN);
            exceptionList.Add(new Structures.ExceptionsStruct { ErrorType = Constants.ErrorTypes.Error, Message = message });
            logger.Error(message);
            return exceptionList;
          }

          var companies = Enumerable.ToList(session.GetAll<Sungero.Parties.ICompany>().Where(x => x.Name == name || 
            (!string.IsNullOrEmpty(tin) && x.TIN == tin && !string.IsNullOrEmpty(trrc) && x.TRRC == trrc) ||
            !string.IsNullOrEmpty(psrn) && x.PSRN == psrn));
          var company = (Enumerable.FirstOrDefault<Sungero.Parties.ICompany>(companies));
          
          if (company != null)
          {
            if (!supplementEntity)
            {
              var message = string.Format("Компания не может быть импортирована. Найден дубль по реквизитам Наименование: \"{0}\", ИНН: {1} + КПП {2}, ОГРН: {3}.", name, tin, trrc, psrn);
              exceptionList.Add(new Structures.ExceptionsStruct {ErrorType = Constants.ErrorTypes.Error, Message = message});
              logger.Error(message);
              return exceptionList;
            }
          }
          else
            company = session.Create<Sungero.Parties.ICompany>();

          company.Name = name;
          company.LegalName = legalName;
          company.HeadCompany = headCompany;
          company.Nonresident = nonresident;
          company.TIN = tin;
          company.TRRC = trrc;
          company.PSRN = psrn;
          company.NCEO = nceo;
          company.NCEA = ncea;
          company.City = city;
          company.Region = region;
          company.LegalAddress = legalAdress;
          company.PostalAddress = postalAdress;
          company.Phones = phones;
          company.Email = email;
          company.Homepage = homepage;
          company.Note = note;
          company.Account = account;
          company.Bank = bank;
          company.Save();
        }
        catch (Exception ex)
        {
          Console.WriteLine(ex.Message);
          exceptionList.Add(new Structures.ExceptionsStruct {ErrorType = Constants.ErrorTypes.Error, Message = ex.Message});
          return exceptionList;
        }
        session.SubmitChanges();
      }
      return exceptionList;
    }
  }
}
