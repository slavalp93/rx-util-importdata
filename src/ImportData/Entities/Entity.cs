using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportData
{
  public class Entity
  {
    public string[] Parameters;
    public Dictionary<string, string> ExtraParameters;
    public int PropertiesCount = 0;

    /// <summary>
    /// Получить наименование число запрашиваемых параметров.
    /// </summary>
    /// <returns>Число запрашиваемых параметров.</returns>
    public virtual int GetPropertiesCount()
    {
      return PropertiesCount;
    }

    /// <summary>
    /// Сохранение сущности в RX.
    /// </summary>
    /// <param name="logger">Логировщик.</param>
    /// <param name="shift">Сдвиг по горизонтали в XLSX документе. Необходим для обработки документов, составленных из элементов разных сущностей.</param>
    /// <returns>Список ошибок.</returns>
    public virtual IEnumerable<Structures.ExceptionsStruct> SaveToRX(NLog.Logger logger, bool supplementEntity, int shift = 0)
    {
      return new List<Structures.ExceptionsStruct>();
    }
  }
}
