using System;
using System.Linq;
using Sungero.Domain.Client;
using System.Security;
using Sungero.Domain.ClientLinqExpressions;
using System.Collections.Generic;
using NLog;
using Sungero.Domain.Shared;
using Sungero.Domain.Client.Deployment;
using Sungero.Metadata.Services;
using CommonLibrary.Dependencies;
using Sungero.Domain.ClientBase;
using Sungero.Presentation;
using NDesk.Options;
using System.IO;
using System.Reflection;

namespace ImportData
{
  class Program
  {
    public static NLog.Logger logger = LogManager.GetCurrentClassLogger();

    /// <summary>
    /// Обновление модулей.
    /// </summary>
    public static void UpdateModules()
    {
      ClientDevelopmentUpdater.Instance.RefreshDevelopment();
      MetadataService.ConfigurationSettingsPaths = new Sungero.Domain.ClientConfigurationSettingsPaths(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
      ClientLazyAssembliesResolver.Instance.LinkToAssembliesFolder(ClientDevelopmentUpdater.Instance.CacheFolder);

      var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
      var developmentDirectory = ClientDevelopmentUpdater.Instance.CacheFolder;

      Dependency.RegisterType<IClientLinqExtensions, ClientLinqExtensionsService>();
      Dependency.RegisterType<IHyperlinkDisplayTextCache, HyperlinkDisplayTextCacheImplementer>();
      Dependency.RegisterType<IHyperlinkEntityCache, HyperlinkEntityCacheImplementer>();

      LoadModules(baseDirectory, "*Client.dll");
      LoadModules(baseDirectory, "*Shared.dll");
      LoadModules(baseDirectory, "*Base.dll");
      LoadModules(developmentDirectory, null);

      EntityFactory.ConfigureUnityContainer();
    }


    /// <summary>
    /// Загрузка модулей.
    /// </summary>
    /// <param name="folderPath">Папка, где находятся модули.</param>
    /// <param name="mask">Маска для поиска файлов.</param>
    private static void LoadModules(string folderPath, string mask)
    {
      if (string.IsNullOrWhiteSpace(mask))
        ModuleManager.Instance.LoadModules(folderPath);
      else
        ModuleManager.Instance.LoadModules(folderPath, mask);
    }

    /// <summary>
    /// Выполнение импорта в соответствии с требуемым действием.
    /// </summary>
    /// <param name="action">Действие.</param>
    /// <param name="xlsxPath">Входной файл.</param>
    /// <param name="extraParameters">Дополнительные параметры.</param>
    /// <param name="logger">Логировщик.</param>
    /// <returns>Соответствующий тип сущности.</returns>
    static void ProcessByAction(string action, string xlsxPath, Dictionary<string, string> extraParameters, NLog.Logger logger)
    {
       
      switch (action)
      {
        case "importcompany":
          logger.Info("Импорт сотрудников");
          logger.Info("-------------");
          EntityProcessor.Procces(typeof(Employee), xlsxPath, Constants.SheetNames.Employees, extraParameters, logger);
          logger.Info("Импорт НОР");
          logger.Info("-------------");
          EntityProcessor.Procces(typeof(BusinessUnit), xlsxPath, Constants.SheetNames.BusinessUnits, extraParameters, logger);
          logger.Info("Импорт подразделений");
          logger.Info("-------------");
          EntityProcessor.Procces(typeof(Department), xlsxPath, Constants.SheetNames.Departments, extraParameters, logger);
          break;
        case "importcompanies":
					EntityProcessor.Procces(typeof(Company), xlsxPath, Constants.SheetNames.Companies, extraParameters, logger);					
          break;				
				
					// Доработка для IMVO - импорт из конкретного xml. Добавить действие в константы!
				case "importcompanies_imvo":
					CompaniesIMVO.Procces(xlsxPath, logger);
					break;

				case "importpersons":
          EntityProcessor.Procces(typeof(Person), xlsxPath, Constants.SheetNames.Persons, extraParameters, logger);
          break;
        case "importcontracts":
          EntityProcessor.Procces(typeof(Contract), xlsxPath, Constants.SheetNames.Contracts, extraParameters, logger);
          break;
        case "importsupagreements":
          EntityProcessor.Procces(typeof(SupAgreement), xlsxPath, Constants.SheetNames.SupAgreements, extraParameters, logger);
          break;
        case "importincomingletters":
          EntityProcessor.Procces(typeof(IncomingLetter), xlsxPath, Constants.SheetNames.IncomingLetters, extraParameters, logger);
          break;
        case "importoutgoingletters":
          EntityProcessor.Procces(typeof(OutgoingLetter), xlsxPath, Constants.SheetNames.OutgoingLetters, extraParameters, logger);
          break;
        case "importorders":
          EntityProcessor.Procces(typeof(Order), xlsxPath, Constants.SheetNames.Orders, extraParameters, logger);
          break;
        default:
          break;
      }
      
    }

    static void Main(string[] args)
    {
      logger.Info("=========================== Process Start ===========================");
      var watch = System.Diagnostics.Stopwatch.StartNew();

			#region Обработка параметров.
			
      var login = string.Empty;
      var password = string.Empty;
      var xlsxPath = string.Empty;
      var action = string.Empty;
			var extraParameters = new Dictionary<string, string>();

      bool isHelp = false;

      var p = new OptionSet() {
        { "n|name=",  "Имя учетной записи DirectumRX.", v => login = v },
        { "p|password=",  "Пароль учетной записи DirectumRX.", v => password = v },
        { "a|action=",  "Действие.", v => action = v },
        { "f|file=",  "Файл с исходными данными.", v => xlsxPath = v },
        { "dr|doc_register_id=",  "Журнал регистрации.", v => extraParameters.Add("doc_register_id", v)},
        { "h|help", "Show this help", v => isHelp = (v != null) },
      };

      try
      {
        p.Parse(args);
      }
      catch (OptionException e)
      {
        Console.WriteLine("Invalid arguments: " + e.Message);
        p.WriteOptionDescriptions(Console.Out);
        return;
      }

      if (isHelp || string.IsNullOrEmpty(login) || string.IsNullOrEmpty(password) || string.IsNullOrEmpty(action)
        || string.IsNullOrEmpty(xlsxPath))
      {
        p.WriteOptionDescriptions(Console.Out);
        return;
      }

      #endregion

      try
      {
        if (!Constants.Actions.dictActions.ContainsKey(action.ToLower()))
        {
          var message = $"Не найдено действие \"{action}\". Введите действие корректно.";
          throw new Exception(message);
        }

        try
        {
          #region Аутентификация.
          System.Security.SecureString sec = new SecureString();
          var passwordCharArray = password.ToCharArray().ToList();
          foreach (var ch in passwordCharArray)
            sec.AppendChar(ch);
          sec.MakeReadOnly();
          var credentials = new UserCredentials(AuthenticationType.UserName, login, sec);
          UserCredentialsManager.Register(credentials, false);
          #endregion

          #region Обновление сборок.
          UpdateModules();
          #endregion

          #region Выполнение импорта сущностей.
          ProcessByAction(action.ToLower(), xlsxPath, extraParameters, logger);

				/*	
          using (var session = new Session())
          {
						//litiko.UAadditions.ICompany company = litiko.UAadditions.Companies.Create();  // using (var session = new Session())
            var company = session.Create<litiko.UAadditions.ICompany>();
						
						// var x = litiko.UAadditions.Companies.Create();
						//company = company as litiko.UAadditions.ICompany;

						company.Name = "Import_Test_Name";
            company.LegalName = "Import_Test_LegalName";
            company.EGRPOUlitiko = "111111111";
            company.Status = Sungero.CoreEntities.DatabookEntry.Status.Active;
            //company.TIN = "123";
            //company.TRRC = trrc;
            //company.PSRN = psrn;
            //company.NCEO = nceo;
            //company.NCEA = ncea;
            //company.City = city;
            //company.Region = region;
            company.LegalAddress = "Import_Test_LegalAddress";
            company.PostalAddress = "Import_Test_PostalAddress";
            company.Phones = "+380951513682";
            company.Email = "slava@litiko.com";                        
            company.Note = "Import_Test_Note"; 
                        
            company.Save();

            session.SubmitChanges();                        
        }
					*/
					#endregion
				}
				catch (Exception ex)
        {
          Console.WriteLine(ex.Message);
          logger.Error(ex.Message);
        }
        finally
        {
          UserCredentialsManager.Unregister();
        }
      }

      catch (Exception ex)
      {
        logger.Error(ex.Message);
      }

      finally
      {
        watch.Stop();
        var elapsedMs = watch.ElapsedMilliseconds;
        logger.Info($"Всего времени затрачено: {elapsedMs} мс");
        logger.Info("=========================== Process Stop ===========================");
      }
    }
  }
}
